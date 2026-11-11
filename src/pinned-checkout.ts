// Checking out an upstream's pinned commit in its clone under .cache, which may predate the pin.
import { execFileSync, spawnSync } from 'node:child_process';

function hasCommit (clone: string, commit: string): boolean {
  return spawnSync('git', [ '-C', clone, 'cat-file', '-e', `${commit}^{commit}` ], { stdio: 'ignore' }).status === 0;
}

/**
 * Checks out `commit` in the clone at `clone`, first fetching from its remote where the clone
 * lacks the commit, and discarding changes to tracked files, such as the patches a runner applies
 * to the upstream sources. Untracked files stay, except where the commit tracks the same path.
 *
 * A full SHA is fetched by itself, so it may be on no branch of the remote; an abbreviated one
 * must be reachable from the remote's branches or tags.
 */
export function checkoutPinned (clone: string, commit: string): void {
  if (!hasCommit(clone, commit)) {
    const refspec = /^[0-9a-f]{40}$/.test(commit) ? [ commit ] : [];
    execFileSync('git', [ '-C', clone, 'fetch', '--quiet', 'origin', ...refspec ], { stdio: 'inherit' });
    if (!hasCommit(clone, commit)) {
      throw new Error(`${clone}: commit ${commit} is not in the clone after fetching from its remote`);
    }
  }
  execFileSync('git', [ '-C', clone, 'checkout', '--quiet', '--force', commit ], { stdio: 'inherit' });
}
