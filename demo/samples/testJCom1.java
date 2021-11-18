import jp.ne.so_net.ga2.no_ji.jcom.*;

public class testJCom1 {
	public static void main(String[] args) throws Exception {
		ReleaseManager rm = new ReleaseManager();
		try {
			System.out.println("EXCELを表示せずに起動...");
			IDispatch xlApp = new IDispatch(rm, "Excel.Application");  // EXCEL本体

			System.out.println("表示されないのは、Visibleがfalseのままだからです。");
			System.out.println("タスクマネージャ等でＥｘｃｅｌが起動しているのが確認できます。");
			System.out.println("\n[Enter]を押してください。終了します");
			System.in.read();
			xlApp.invoke("Quit", null);

			xlApp.release();
			System.out.println("ご静聴、ありがとうございました。");
		}
		catch(Exception e) { e.printStackTrace(); }
		finally { rm.release(); }
	}
}
