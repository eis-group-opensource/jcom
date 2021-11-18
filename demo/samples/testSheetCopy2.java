import jp.ne.so_net.ga2.no_ji.jcom.*;
import jp.ne.so_net.ga2.no_ji.jcom.excel8.*;

public class testSheetCopy2 {
	public static void main(String[] args) throws Exception {
		ReleaseManager rm = new ReleaseManager();
		try {
			System.out.println("EXCELを起動中...");
			ExcelApplication xlApp = new ExcelApplication(rm);
			xlApp.Visible(true);
			ExcelWorkbooks xlBooks = xlApp.Workbooks();
			ExcelWorkbook xlBook = xlBooks.Add();	// create new book
			ExcelWorksheet xlSheet = xlApp.ActiveSheet();

			// set string to cell A1 .
			System.out.println("セルA1に文字列をセット");
			ExcelRange xlRange = xlSheet.Range("A1");
			xlRange.Value("JCom (^o^)");

			// copy cell from A1 to B2 .
			// セルをコピーしてみる。単一セル
			System.out.println("セルA1の内容をB1にコピー");
			xlRange.Copy(xlSheet.Range("B2"));

			// copy cells from A1:B2 to C1:D2 .
			// セルをコピーしてみる。複数セル A1:B2を C1:D2へコピー
			System.out.println("セルA1:B2の内容をC1:D2へコピー");
			ExcelRange xlRangeA1B2 = xlSheet.Range("A1:B2");
			xlRangeA1B2.Copy(xlSheet.Range("C1"));

			// copy sheet.
			// シートをコピーしてみる
			System.out.println("シートをコピー");
			xlSheet.Copy(null, xlSheet);

			System.out.println("Hit [Enter] key to exit.");
			System.in.read();

			// quit.
			// 終了させる。
			xlBook.Close(false, null, false);
			xlApp.Quit();
		}
		catch(Exception e) { e.printStackTrace(); }
		finally { rm.release(); }
	}
}
