import jp.ne.so_net.ga2.no_ji.jcom.*;
import jp.ne.so_net.ga2.no_ji.jcom.excel8.*;

class testEnum {

	public static void main(String[] args) throws Exception {
		ReleaseManager rm = new ReleaseManager();
		try {
			System.out.println("EXCELを起動中...");
			// すでに立ち上がっていると、新しいウィンドウで開く。
			ExcelApplication excel = new ExcelApplication(rm);
			excel.Visible(true);

			ExcelWorkbooks xlBooks = excel.Workbooks();
			ExcelWorkbook xlBook = xlBooks.Add();	// 新しいブックを作成
			ExcelWorksheets xlSheets = xlBook.Worksheets();

			// シートからIEnumVARIANTを生成。Enumっぽい操作
			System.out.println("すべてのシート名を列挙してみる");
			IEnumVARIANT enum = xlSheets.NewEnum();
			for(;;) {
				IDispatch disp = (IDispatch)enum.next();
				if(disp==null) break;
				ExcelWorksheet xlSheet = new ExcelWorksheet(disp);
				System.out.println(""+xlSheet.Name());
			}

			System.out.println("別の方法で同じことをやってみる");
			enum.reset();	// 最初から
			Object[] ary = enum.next(100);	// 最大１００個のオブジェクトを取得
			for(int i=0; i<ary.length; i++) {
				ExcelWorksheet xlSheet = new ExcelWorksheet((IDispatch)ary[i]);
				System.out.println(""+xlSheet.Name());
			}

			System.out.println("[Enter]を押してください。終了します");
			System.in.read();

			xlBook.Close(false,null,false);
			excel.Quit();

			System.out.println("ご静聴、ありがとうございました。");
		}
		catch(Exception e) { e.printStackTrace(); }
		finally { rm.release(); }
	}
}
