// Shim for replay/spreadsheetgear.ts: opens one fixture with SpreadsheetGear's own xlsx reader,
// forces a full recalculation and writes the values of the requested cells as JSON.
//
//   SgEvaluate <fixture.xlsx> <targets.json> <values.json>
//
// targets.json: [{"sheet": "Sheet1", "ref": "A1"}, ...]
// values.json:  {"Sheet1!A1": <value>, ...} where a value is a number, a string, a boolean, an
//               error string such as "#DIV/0!", or null for an empty cell or a missing sheet.
//
// It then prints the managed heap's retained heap and peak live heap (CONTEXT.md), in MiB, as
//   #memory \t M \t <retained heap> - <peak live heap>
// with `-` for the retained external memory, which .NET does not account for.
//
// SPREADSHEETGEAR_LICENSE, when set, is passed to Factory.SetSignedLicense (the 30-day trial or a
// paid licence). Without it SpreadsheetGear runs under its no-cost licence, which refuses
// workbooks over 1,000 rows x 100 columns x 10 worksheets; such a fixture fails to load.
using System.Diagnostics.Tracing;
using System.Globalization;
using System.Text.Json;
using SpreadsheetGear;

using var fullCollections = new FullCollectionWatch();

var license = Environment.GetEnvironmentVariable("SPREADSHEETGEAR_LICENSE");
if (!string.IsNullOrWhiteSpace(license)) Factory.SetSignedLicense(license.Trim());

var targets = JsonSerializer.Deserialize<List<Dictionary<string, string>>>(File.ReadAllText(args[1])) ?? [];
var set = Factory.GetWorkbookSet(new CultureInfo("en-US"));
var workbook = set.Workbooks.Open(args[0]);
set.CalculateFull();

var errors = new Dictionary<string, string> {
    ["Div0"] = "#DIV/0!", ["NA"] = "#N/A", ["Name"] = "#NAME?", ["Null"] = "#NULL!",
    ["Num"] = "#NUM!", ["Ref"] = "#REF!", ["Value"] = "#VALUE!", ["Spill"] = "#SPILL!",
    ["Calc"] = "#CALC!", ["GettingData"] = "#GETTING_DATA", ["Field"] = "#FIELD!",
    ["Blocked"] = "#BLOCKED!", ["Connect"] = "#CONNECT!", ["Busy"] = "#BUSY!", ["Unknown"] = "#UNKNOWN!",
};

var output = new Dictionary<string, object?>();
foreach (var target in targets) {
    var sheetName = target["sheet"];
    var key = $"{sheetName}!{target["ref"]}";
    var sheet = workbook.Worksheets[sheetName];
    if (sheet == null) { output[key] = null; continue; }
    object? value = sheet.Cells[target["ref"]].Value;
    output[key] = value switch {
        null => null,
        double d => double.IsFinite(d) ? d : "#NUM!",
        string s => s.Length == 0 ? null : s,
        bool b => b,
        ValueError e => errors.TryGetValue(e.ToString(), out var text) ? text : $"#{e.ToString().ToUpperInvariant()}!",
        _ => Convert.ToString(value, CultureInfo.InvariantCulture),
    };
}
File.WriteAllText(args[2], JsonSerializer.Serialize(output));

// Measured while the workbook is still reachable. The forced collection's event may arrive later,
// so its heap is counted directly.
GC.Collect();
GC.WaitForPendingFinalizers();
GC.Collect();
var retainedHeap = FullCollectionWatch.LiveHeapAfter(GCKind.FullBlocking);
var peakLiveHeap = Math.Max(fullCollections.Peak, retainedHeap);
GC.KeepAlive(workbook);
static long Mib (long bytes) => (long)Math.Round(bytes / 1048576.0);
Console.Write($"#memory\tM\t{Mib(retainedHeap)} - {Mib(peakLiveHeap)}\n");
return 0;

// Tracks the most managed heap in use after a full collection, from the runtime's GC events. A gen2
// collection, blocking or background, marks the whole heap; gen0 and gen1 collections leave garbage
// in gen2, so they are not counted. An event arrives on another thread after its collection, and
// GC.GetGCMemoryInfo then gives the latest full collection, which may be a later one: a collection
// can be missed, but not misread.
sealed class FullCollectionWatch : EventListener {
    long peak;

    public long Peak => Interlocked.Read(ref peak);

    // The heap in use after the latest collection of a kind: the generations' sizes, LOH and POH
    // included, less their free space. 0 before any.
    public static long LiveHeapAfter (GCKind kind) {
        long live = 0;
        foreach (var generation in GC.GetGCMemoryInfo(kind).GenerationInfo) {
            live += generation.SizeAfterBytes - generation.FragmentationAfterBytes;
        }
        return live;
    }

    protected override void OnEventSourceCreated (EventSource source) {
        // Keyword 0x1 is the runtime's GC keyword.
        if (source.Name == "Microsoft-Windows-DotNETRuntime") EnableEvents(source, EventLevel.Informational, (EventKeywords)0x1);
    }

    protected override void OnEventWritten (EventWrittenEventArgs e) {
        if (e.EventName?.StartsWith("GCEnd") != true) return;
        var live = Math.Max(LiveHeapAfter(GCKind.FullBlocking), LiveHeapAfter(GCKind.Background));
        long seen;
        while (live > (seen = Interlocked.Read(ref peak)) && Interlocked.CompareExchange(ref peak, live, seen) != seen) { }
    }
}
