import jp.ne.so_net.ga2.no_ji.jcom.*;

public class testSheetCopy {
	public static void main(String[] args) throws Exception {
		ReleaseManager rm = new ReleaseManager();
		try {
			System.out.println("EXCELを起動中...");
			IDispatch xlApp = new IDispatch(rm, "Excel.Application");
			xlApp.put("Visible", new Boolean(true));
			IDispatch xlBooks = (IDispatch)xlApp.get("Workbooks");
			IDispatch xlBook = (IDispatch)xlBooks.method("Add",null);	// create new book
			IDispatch xlSheet = (IDispatch)xlApp.get("ActiveSheet");

			// set string to cell A1 .
			System.out.println("セルA1に文字列をセット");
			Object[] arglist = new Object[1];
			arglist[0] = "A1";
			IDispatch xlRange = (IDispatch)xlSheet.get("Range",arglist);
			xlRange.put("Value","JCom (^o^)");

			// copy cell from A1 to B2 .
			// セルをコピーしてみる。単一セル
			System.out.println("セルA1の内容をB1にコピー");
			Object[] copyargs = new Object[1];
			copyargs[0] = (IDispatch)xlSheet.get("Range", new Object[] {"B2"});
			xlRange.method("Copy", copyargs);

			// copy cells from A1:B2 to C1:D2 .
			// セルをコピーしてみる。複数セル A1:B2を C1:D2へコピー
			System.out.println("セルA1:B2の内容をC1:D2へコピー");
			IDispatch xlRangeA1B2 = (IDispatch)xlSheet.get("Range",new Object[] {"A1:B2"});
			copyargs = new Object[1];
			copyargs[0] = (IDispatch)xlSheet.get("Range", new Object[] {"C1"});
			xlRangeA1B2.method("Copy", copyargs);

			// copy sheet.
			// シートをコピーしてみる
			System.out.println("シートをコピー");
			copyargs = new Object[2];
			copyargs[0] = null;
			copyargs[1] = xlSheet;
			xlSheet.method("Copy", copyargs);

			System.out.println("Hit [Enter] key to exit.");
			System.in.read();

			// quit.
			// 終了させる。
			Object[] arglist3 = new Object[3];
			arglist3[0] = new Boolean(false);
			arglist3[1] = null;
			arglist3[2] = new Boolean(false);
			xlBook.method("Close", arglist3);
			xlApp.method("Quit", null);
		}
		catch(Exception e) { e.printStackTrace(); }
		finally { rm.release(); }
	}
}
