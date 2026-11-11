"""Answers each request line on stdin with one line on stdout.

- `tree <pid>` is answered with `<pid>=<footprint>`, the footprint in bytes, for the process and
  each of its descendants, the process first, or with an empty line if the process does not exist.
- `pressure` is answered with the system's memory pressure level: `normal`, `warn`, `critical`,
  or `unknown` where the system does not say.
- `available` is answered with the memory available to processes in bytes, or `-` where the system
  does not say: on macOS the kernel's free percentage (`kern.memorystatus_level`, which
  `memory_pressure -Q` prints) of the installed memory, on Linux the MemAvailable of /proc/meminfo.

Any other request stops the helper with an error.

The physical footprint (CONTEXT.md) is read from the kernel's per-process ledger, which top and
Activity Monitor read too (MEMORY ACCOUNTING in `man footprint`). Linux has no such ledger: there
it is the anonymous resident memory and the swap of /proc/<pid>/status. A process that cannot be
read, for example one that has exited, answers `-`.

Node cannot call proc_pid_rusage without a native addon, so src/memory.ts runs this script and
keeps it running for a watch. It finds the descendants too, through libproc or /proc, and reads
the memory pressure through sysctl or /proc, so that a sample starts no process.

On macOS the memory pressure level is the kernel's. Linux has none: there it is derived from the
pressure stall information of /proc/pressure/memory, the share of the last 10 seconds in which
tasks stalled waiting for memory: `warn` when some task stalled at least 10% of that time,
`critical` when all tasks stalled together at least 10% of it. A kernel without that file, or with
it disabled, answers `unknown`.
"""
import ctypes
import os
import sys

RUSAGE_INFO_V0 = 0
PROC_PPID_ONLY = 6
# kern.memorystatus_vm_pressure_level: the NOTE_MEMORYSTATUS_PRESSURE_* flags of xnu's
# <sys/event_private.h>. The kernel reports its urgent level, between warn and critical, as warn.
DARWIN_PRESSURE_LEVELS = {1: 'normal', 2: 'warn', 4: 'critical'}
# The percentages of stalled time, over the last 10 seconds, at which Linux pressure is elevated.
PSI_WARN_SOME_AVG10 = 10.0
PSI_CRITICAL_FULL_AVG10 = 10.0


class RusageInfoV0(ctypes.Structure):
    """struct rusage_info_v0 of <sys/resource.h>, the smallest flavour with ri_phys_footprint."""
    _fields_ = [
        ('ri_uuid', ctypes.c_uint8 * 16),
        ('ri_user_time', ctypes.c_uint64),
        ('ri_system_time', ctypes.c_uint64),
        ('ri_pkg_idle_wkups', ctypes.c_uint64),
        ('ri_interrupt_wkups', ctypes.c_uint64),
        ('ri_pageins', ctypes.c_uint64),
        ('ri_wired_size', ctypes.c_uint64),
        ('ri_resident_size', ctypes.c_uint64),
        ('ri_phys_footprint', ctypes.c_uint64),
        ('ri_proc_start_abstime', ctypes.c_uint64),
        ('ri_proc_exit_abstime', ctypes.c_uint64),
    ]


if sys.platform == 'darwin':
    libc = ctypes.CDLL('/usr/lib/libSystem.B.dylib', use_errno=True)
    libc.proc_pid_rusage.argtypes = [ctypes.c_int, ctypes.c_int, ctypes.c_void_p]
    libc.proc_pid_rusage.restype = ctypes.c_int
    libc.proc_listpids.argtypes = [ctypes.c_uint32, ctypes.c_uint32, ctypes.c_void_p, ctypes.c_int]
    libc.proc_listpids.restype = ctypes.c_int
    libc.sysctlbyname.argtypes = [ctypes.c_char_p, ctypes.c_void_p, ctypes.POINTER(ctypes.c_size_t), ctypes.c_void_p, ctypes.c_size_t]
    libc.sysctlbyname.restype = ctypes.c_int

    def pressure():
        level = ctypes.c_int(0)
        size = ctypes.c_size_t(ctypes.sizeof(level))
        if libc.sysctlbyname(b'kern.memorystatus_vm_pressure_level', ctypes.byref(level), ctypes.byref(size), None, 0) != 0:
            return 'unknown'
        return DARWIN_PRESSURE_LEVELS.get(level.value, 'unknown')

    def sysctl_int(name, ctype):
        value = ctype(0)
        size = ctypes.c_size_t(ctypes.sizeof(value))
        if libc.sysctlbyname(name, ctypes.byref(value), ctypes.byref(size), None, 0) != 0:
            return None
        return value.value

    def available():
        percent = sysctl_int(b'kern.memorystatus_level', ctypes.c_int)
        installed = sysctl_int(b'hw.memsize', ctypes.c_uint64)
        return None if percent is None or installed is None else installed * percent // 100

    def phys_footprint(pid):
        info = RusageInfoV0()
        if libc.proc_pid_rusage(pid, RUSAGE_INFO_V0, ctypes.byref(info)) != 0:
            return None
        return info.ri_phys_footprint

    def children_of():
        """A function from a process ID to the IDs of its children."""
        def children(pid):
            # The kernel answers with the bytes it filled; a full buffer may have been cut short.
            size = 64
            while True:
                buffer = (ctypes.c_int * size)()
                filled = libc.proc_listpids(PROC_PPID_ONLY, pid, buffer, ctypes.sizeof(buffer))
                if filled <= 0:
                    return []
                count = filled // ctypes.sizeof(ctypes.c_int)
                if count < size:
                    return [p for p in buffer[:count] if p > 0]
                size *= 4
        return children
else:
    def phys_footprint(pid):
        try:
            with open(f'/proc/{pid}/status') as status:
                fields = dict(line.split(':', 1) for line in status if ':' in line)
            return sum(int(fields[name].split()[0]) * 1024 for name in ('RssAnon', 'VmSwap'))
        except (OSError, KeyError, ValueError):
            return None

    def children_of():
        """A function from a process ID to the IDs of its children, as of one scan of /proc."""
        table = {}
        for entry in os.listdir('/proc'):
            if not entry.isdigit():
                continue
            try:
                with open(f'/proc/{entry}/stat') as stat:
                    # The command name, in parentheses, may itself contain spaces and parentheses.
                    ppid = int(stat.read().rsplit(')', 1)[1].split()[1])
            except (OSError, IndexError, ValueError):
                continue
            table.setdefault(ppid, []).append(int(entry))
        return lambda pid: table.get(pid, [])

    def pressure():
        try:
            with open('/proc/pressure/memory') as psi:
                return psi_level(psi.read())
        except (OSError, KeyError, ValueError):
            return 'unknown'

    def available():
        try:
            with open('/proc/meminfo') as meminfo:
                return meminfo_available(meminfo.read())
        except (OSError, KeyError, ValueError, IndexError):
            return None


def meminfo_available(text):
    """The MemAvailable of the text of /proc/meminfo, in bytes: lines like `MemAvailable: 123 kB`."""
    fields = dict(line.split(':', 1) for line in text.splitlines() if ':' in line)
    return int(fields['MemAvailable'].split()[0]) * 1024


def psi_level(text):
    """The pressure level of the text of /proc/pressure/memory: lines like `some avg10=0.00 ...`."""
    avg10 = {}
    for line in text.splitlines():
        kind, *fields = line.split()
        avg10[kind] = float(dict(field.split('=', 1) for field in fields)['avg10'])
    if avg10.get('full', 0) >= PSI_CRITICAL_FULL_AVG10:
        return 'critical'
    if avg10['some'] >= PSI_WARN_SOME_AVG10:
        return 'warn'
    return 'normal'


def exists(pid):
    try:
        os.kill(pid, 0)
    except ProcessLookupError:
        return False
    except PermissionError:
        return True
    return True


def tree(pid):
    """A process and all its descendants, the process first, or an empty list if it does not exist."""
    if not exists(pid):
        return []
    children = children_of()
    found = []
    pending = [pid]
    seen = set()
    while pending:
        p = pending.pop()
        if p in seen:
            continue
        seen.add(p)
        found.append(p)
        pending.extend(children(p))
    return found


def answer(footprint):
    return '-' if footprint is None else str(footprint)


if __name__ == '__main__':
    for line in sys.stdin:
        words = line.split()
        if words[:1] == ['pressure']:
            print(pressure(), flush=True)
        elif words[:1] == ['available']:
            print(answer(available()), flush=True)
        elif words[:1] == ['tree']:
            print(' '.join(f'{p}={answer(phys_footprint(p))}' for p in tree(int(words[1]))), flush=True)
        else:
            sys.exit(f'unknown request: {line.strip()}')
