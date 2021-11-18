import jp.ne.so_net.ga2.no_ji.jcom.excel8.*;
import jp.ne.so_net.ga2.no_ji.jcom.*;
import java.io.File;
import java.util.Date;

/**
	スレッド単位にＣＯＭを実行するサンプル。
	スレッドごとにReleaseManagerを割り当てている。
*/
class testExcelSvr extends Thread implements Runnable {
	
	// スレッドを ? 個管理するクラス
	static final int THREAD_NUM = 1;	// なぜか１以外では希にエラーが起こる。原因不明
	static final int REQUESTS = 100;
	static testExcelSvr[] list;

	public synchronized void run() {
		ReleaseManager rm = new ReleaseManager();
		try {
			System.out.println("EXCELを起動中...");
			// すでに立ち上がっていると、新しいウィンドウで開く。
			ExcelApplication excel = new ExcelApplication(rm);
			excel.Visible(true);

			ExcelWorkbooks xlBooks = excel.Workbooks();
			ExcelWorkbook xlBook = xlBooks.Add();	// 新しいブックを作成
			
			// すべてのファイルを列挙してみる
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
//			xlRange.Item(1,8).Value("隠しファイル"); // jdk1.2以降

			File path = new File("./");
			String[] filenames = path.list();
			for(int i=0; i<filenames.length; i++) {
				//System.out.println(files[i]);
				File file = new File(filenames[i]);
				xlRange.Item(i+2,1).Value( file.getName() );				// ファイル名パス無し
				xlRange.Item(i+2,2).Value( (int)file.length() );			// ファイルサイズ
				xlRange.Item(i+2,3).Value( new Date(file.lastModified()) );	// 最終更新日時
				xlRange.Item(i+2,4).Value( file.isDirectory()?"Yes":"No" );	// ディレクトリか？
				xlRange.Item(i+2,5).Value( file.isFile()?"Yes":"No" );		// ファイルか？
				xlRange.Item(i+2,6).Value( file.canRead()?"Yes":"No" );		// 読み取り可か？
				xlRange.Item(i+2,7).Value( file.canWrite()?"Yes":"No" );	// 書き込み可か？
//				xlRange.Item(i+2,8).Value( file.isHidden()?"Yes":"No" );	// 隠しファイルか？
			}

			xlRange.Item(filenames.length+2,1).Value("合計");
			xlRange.Item(filenames.length+2,2).Formula("=Sum(B2:B"+(filenames.length+1)+")");
			xlRange.Columns().AutoFit();	// 横幅をフィットさせる

			xlBook.Close(false,null,false);
			excel.Quit();
		}
		catch(Exception e) { e.printStackTrace(); }
		finally { rm.release(); }
		// 自分のスレッド情報を消す
		for(int i=0; i<THREAD_NUM; i++) {
			if(list[i] == this) {
				list[i] = null;
				break;
			}
		}
		System.out.println("スレッド終了");
	}
	
	
	public static void main(String[] args) {
		System.out.println("*** サーバーアプリを想定したサンプル ***");
		System.out.println(""+REQUESTS+"個の要求を"+THREAD_NUM+"個のスレッドで同時に非同期に処理します。");
		System.out.println("すべての処理を終えるのに、しばらく時間がかかります。");
		//スレッドリストを初期化
		list = new testExcelSvr[THREAD_NUM];
		for(int i=0; i<THREAD_NUM; i++) {
			list[i] = null;
		}
		for(int count=0; count<REQUESTS; count++) {
			// 空きがあったら、スレッドを開始
			for(int i=0; i<THREAD_NUM; i++) {
				if(list[i]==null) {
					System.out.println("スレッド生成");
					list[i] = new testExcelSvr();
					list[i].start();
					break;	// 1回のループにつき１つ生成
				}
			}
			try{sleep(1000);}catch(InterruptedException e){}	// 1sec
		}
		// すべてのスレッドが終わるまで待つ。
		System.out.println("すべてのスレッドが終わるまで待つ。");
		for(;;) {
			boolean endflag = true;
			for(int i=0; i<THREAD_NUM; i++) {
				if(list[i]!=null) {
					endflag = false;
				}
			}
			if(endflag) break;
			try{sleep(1000);}catch(InterruptedException e){}	// 1sec
		}
	}

}
/*
サーバアプリケーションのサンプル。
サーバアプリケーションはこのような、マルチスレッドや排他処理を
考慮する必要があります。排他処理は、JComのCOMを呼ぶインターフェース
部分で行っています。Excelを複数同時に立ちあげることは可能ですが、
私が負荷の高いテストしたところ、たまにCOMの呼び出しが失敗しました。
ここでは、同時に実行できないよう、THREAD_NUMを１にしています。
ここを変更すれば、同時に複数の処理を並行して行います。
*/

