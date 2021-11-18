import jp.ne.so_net.ga2.no_ji.jcom.*;

/*
	ReleaseManagerを使わない場合のサンプル
	確保されたＣＯＭオブジェクトに対して、必ずrelease()を呼べば
	プロセスを正常に終了することができます。しかし、ソースが複雑に
	なりますし、どこでrelease()を呼べばよいのかがわかりにくいです。
	ReleaseManagerはそれを管理し、適切に処理してくれます。
	ReleaseManagerはあらゆるケースを想定して設計されているので、
	このソースを参考にすることはないでしょう。
*/
public class testNoReleaseManager {
	public static void main(String[] args) throws Exception {
		try {
			System.out.println("EXCELを起動中...");
			IDispatch xlApp = new IDispatch(null, "Excel.Application");  // EXCEL本体
			xlApp.put("Visible", new Boolean(true));	// 'デフォルトはFalse(表示しない)
			IDispatch xlBooks = (IDispatch)xlApp.get("Workbooks");
			IDispatch xlBook = (IDispatch)xlBooks.invoke("Add",null);	// 新しいブックを作成
			IDispatch xlSheet = (IDispatch)xlApp.get("ActiveSheet");

			System.out.println("セルA1に文字列をセット");
			Object[] arglist = new Object[1];
			arglist[0] = "A1";
			IDispatch xlRange = (IDispatch)xlSheet.get("Range",arglist);
			xlRange.put("Value","JComすごいぞ！(^o^)");
			xlRange.release();		// @@@

			System.out.println("[Enter]を押してください。終了します");
			System.in.read();

			Object[] arglist3 = new Object[3];
			arglist3[0] = new Boolean(false);
			arglist3[1] = null;
			arglist3[2] = new Boolean(false);
			xlBook.invoke("Close", arglist3);
			xlApp.invoke("Quit", null);
			xlSheet.release();		// @@@
			xlBook.release();		// @@@
			xlBooks.release();		// @@@
			xlApp.release();		// @@@
			System.out.println("ご静聴、ありがとうございました。");
		}
		catch(Exception e) { e.printStackTrace(); }
	}
}
