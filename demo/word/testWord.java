import jp.ne.so_net.ga2.no_ji.jcom.*;

/**
	ワードのサンプル
	2001.07.04
*/
public class testWord {

	public static void main(String[] args) throws Exception {
		ReleaseManager rm = new ReleaseManager();
		try {
			System.out.println("Wordを起動中...");
			IDispatch wdApp = new IDispatch(rm, "Word.Application");  // EXCEL本体
			wdApp.put("Visible", new Boolean(true));	// 'デフォルトはFalse(表示しない)

			IDispatch wdDocuments = (IDispatch)wdApp.get("Documents");
			Object[] arglist1 = new Object[1];
			String userdir = System.getProperty("user.dir");	// user.dir="E:\USERS\java\test"
			arglist1[0] = userdir+"\\jcom.doc";
			IDispatch wdDocument = (IDispatch)wdDocuments.method("Open", arglist1);
			String fullname = (String)wdDocument.get("FullName");
			System.out.println("fullname="+fullname);

			// 単語を見てみる
			IDispatch wdWords = (IDispatch)wdDocument.get("Words");
			int word_count = ((Integer)wdWords.get("Count")).intValue();
			for(int i=0; i<word_count; i++) {
				Object[] index = new Object[1];
				index[0] = new Integer(i+1);		// COMコレクションは１から始まる
				IDispatch wdWord = (IDispatch)wdWords.method("Item", index);
				String text = (String)wdWord.get("Text");
				System.out.println(text);
			}

			// 表を見てみる
			IDispatch wdTables = (IDispatch)wdDocument.get("Tables");
			System.out.println(wdTables);
			int table_count = ((Integer)wdTables.get("Count")).intValue();
			System.out.println("表の数="+table_count);
			for(int i=0; i<table_count; i++) {
				Object[] index = new Object[1];
				index[0] = new Integer(i+1);		// COMコレクションは１から始まる
				IDispatch wdTable = (IDispatch)wdTables.method("Item", index);
				System.out.println(""+i+"="+wdTable);
				// 一応、表は取れるけれど・・・
			}

			// プリンタに出力
			//wdDocument.method("PrintOut", null);	// 動作未確認

			System.out.println("３秒後に終了します。");
			Thread.sleep(3000);	// 3sec
			wdApp.method("Quit", null);
			System.out.println("ご静聴、ありがとうございました。");
		}
		catch(Exception e) { e.printStackTrace(); }
		finally { rm.release(); }
	}
}
