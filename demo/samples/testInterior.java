import jp.ne.so_net.ga2.no_ji.jcom.excel8.*;
import jp.ne.so_net.ga2.no_ji.jcom.*;
import java.io.File;
import java.util.Date;

/* Excel用ラッパを使った、JComのサンプルプログラム 
	背景色、ボーダーのテスト
*/
class testInterior {
	public static void main(String[] args) throws Exception {
		ReleaseManager rm = new ReleaseManager();
		try {
			System.out.println("EXCELを起動中...");
			// すでに立ち上がっていると、新しいウィンドウで開く。
			ExcelApplication excel = new ExcelApplication(rm);
			excel.Visible(true);
			// 色んな情報を表示
			System.out.println("Version="+excel.Version());
			System.out.println("UserName="+excel.UserName());
			System.out.println("Caption="+excel.Caption());
			System.out.println("Value="+excel.Value());

			ExcelWorkbooks xlBooks = excel.Workbooks();
			ExcelWorkbook xlBook = xlBooks.Add();	// 新しいブックを作成
			
			// すべてのファイルを列挙してみる
			System.out.println("現在のディレクトリのファイルをセルに設定");
			ExcelWorksheets xlSheets = xlBook.Worksheets();
			ExcelWorksheet xlSheet = xlSheets.Item(1);
			ExcelRange xlRange = xlSheet.Cells();

			xlRange.Item(1,1).Value("ファイル名" );
			xlRange.Item(1,2).Value("サイズ" );
			xlRange.Item(1,3).Value("最終更新日時");
			xlRange.Item(1,4).Value("ディレクトリ");
			xlRange.Item(1,5).Value("ファイル");
			xlRange.Item(1,6).Value("読み込み可");
			xlRange.Item(1,7).Value("書き込み可");
//			xlRange.Item(1,8).Value("隠しファイル");

			// 低レベルインターフェースを使って、サポートされていないオブジェクトにアクセスしてみる。
			// セルの背景色に色を設定。
			IDispatch interior = (IDispatch)xlRange.Item(1,1).get("Interior");
			interior.put("Color", new Integer(0xFFFF00));	// GGBBRR cyan
			// 罫線を設定。
			IDispatch borders = (IDispatch)xlRange.Item(1,1).get("Borders");
			Object[] border_args = new Integer[1];
			border_args[0] = new Integer(9);
			IDispatch border = (IDispatch)borders.get("Item",border_args);	// XlBordersIndex.xlEdgeBottom = 9
			border.put("LineStyle", new Integer(1));	// XlLineStyle.xlContinuous = 1


			File path = new File("./");
			String[] filenames = path.list();
			for(int i=0; i<filenames.length; i++) {
				File file = new File(filenames[i]);
				System.out.println(file);
				xlRange.Item(i+2,1).Value( file.getName() );				// ファイル名パス無し
				xlRange.Item(i+2,2).Value( (int)file.length() );			// ファイルサイズ
				xlRange.Item(i+2,3).Value( new Date(file.lastModified()) );	// 最終更新日時
				xlRange.Item(i+2,4).Value( file.isDirectory()?"Yes":"No" );	// ディレクトリか？
				xlRange.Item(i+2,5).Value( file.isFile()?"Yes":"No" );		// ファイルか？
				xlRange.Item(i+2,6).Value( file.canRead()?"Yes":"No" );		// 読み取り可か？
				xlRange.Item(i+2,7).Value( file.canWrite()?"Yes":"No" );	// 書き込み可か？
//				xlRange.Item(i+2,8).Value( file.isHidden()?"Yes":"No" );	// 隠しファイルか？ (jdk1.2以降)
			}
			String expression = "=Sum(B2:B"+(filenames.length+1)+")";
			System.out.println("数式を埋め込み、ファイルサイズの合計を求める "+expression);
			xlRange.Item(filenames.length+2,1).Value("合計");
			xlRange.Item(filenames.length+2,2).Formula(expression);
			xlRange.Columns().AutoFit();	// 横幅をフィットさせる

			// プリンタに出力する場合はコメントをはずしてください。
			// デフォルトのプリンタに出力されます。
//			System.out.println("プリンタに印刷します。");
//			xlSheet.PrintOut();

			// ファイルに保存する場合はコメントを外してください。
			// ディレクトリを指定しない場合は、(My Documents)に保存されます。
//			System.out.println("ファイルに保存します。(My Documents)\\testExcel.xls");
//			xlBook.SaveAs("testExcel.xls");

			System.out.println("10秒後に終了します");
			Thread.sleep(10*1000);

			xlBook.Close(false,null,false);
			excel.Quit();

			System.out.println("ご静聴、ありがとうございました。");
		}
		catch(Exception e) { e.printStackTrace(); }
		finally { rm.release(); }
	}
}
