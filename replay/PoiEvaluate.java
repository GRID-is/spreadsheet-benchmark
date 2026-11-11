// Shim for replay/poi.ts: evaluates the requested case cells of one workbook
// with Apache POI's FormulaEvaluator and prints one line per target:
//   <sheet>!<ref> \t <kind> \t <payload>
// kinds: N number, S string (\\ \t \n escaped), B boolean, E Excel error
// string, U unsupported function (reported as #NAME?), X evaluation crash,
// NULL empty. POI parses the xlsx itself; evaluation is per target cell so a
// NotImplementedException poisons one case, not the file. A last line gives the
// JVM's heap figures (replay/JvmHeap.java).
import java.io.File;
import java.nio.file.Files;
import java.nio.file.Paths;
import org.apache.poi.ss.formula.eval.NotImplementedException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaError;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellReference;

public class PoiEvaluate {
    static String escape (String s) {
        return s.replace("\\", "\\\\").replace("\t", "\\t").replace("\n", "\\n").replace("\r", "\\r");
    }

    public static void main (String[] args) throws Exception {
        JvmHeap.watchFullCollections();
        try (Workbook wb = WorkbookFactory.create(new File(args[0]), null, true)) {
            FormulaEvaluator ev = wb.getCreationHelper().createFormulaEvaluator();
            ev.setIgnoreMissingWorkbooks(true);
            StringBuilder out = new StringBuilder();
            for (String line : Files.readAllLines(Paths.get(args[1]))) {
                int tab = line.indexOf('\t');
                if (tab < 0) continue;
                String sheetName = line.substring(0, tab);
                String ref = line.substring(tab + 1);
                String key = sheetName + "!" + ref;
                Sheet sheet = wb.getSheet(sheetName);
                if (sheet == null) { out.append(key).append("\tNULL\t\n"); continue; }
                CellReference cr = new CellReference(ref);
                Row row = sheet.getRow(cr.getRow());
                Cell cell = row == null ? null : row.getCell(cr.getCol());
                if (cell == null) { out.append(key).append("\tNULL\t\n"); continue; }
                try {
                    CellValue v = ev.evaluate(cell);
                    if (v == null) { out.append(key).append("\tNULL\t\n"); continue; }
                    switch (v.getCellType()) {
                        case NUMERIC: out.append(key).append("\tN\t").append(v.getNumberValue()).append('\n'); break;
                        case STRING: out.append(key).append("\tS\t").append(escape(v.getStringValue())).append('\n'); break;
                        case BOOLEAN: out.append(key).append("\tB\t").append(v.getBooleanValue()).append('\n'); break;
                        case ERROR: out.append(key).append("\tE\t").append(FormulaError.forInt(v.getErrorValue()).getString()).append('\n'); break;
                        default: out.append(key).append("\tNULL\t\n");
                    }
                }
                catch (NotImplementedException e) {
                    out.append(key).append("\tU\t\n");
                }
                catch (Exception | StackOverflowError e) {
                    out.append(key).append("\tX\t").append(escape(String.valueOf(e.getMessage()))).append('\n');
                }
            }
            // Measured while wb is still reachable: try-with-resources closes it after this block.
            out.append(JvmHeap.memoryLine());
            System.out.print(out);
        }
    }
}
