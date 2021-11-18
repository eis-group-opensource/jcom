import jp.ne.so_net.ga2.no_ji.jcom.*;

public class testError {
	public static void main(String[] args) throws Exception {
		ReleaseManager rm = new ReleaseManager();
		try {
			System.out.println("EXCELを起動中...");
			IDispatch xlApp = new IDispatch(rm, "Excel.Application");  // EXCEL本体
			xlApp.put("Visible", new Boolean(true));	// 'デフォルトはFalse(表示しない)

			IDispatch xlBooks = (IDispatch)xlApp.get("Workbooks");
			
			IDispatch xlBook = (IDispatch)xlBooks.method("Add",null);	// 新しいブックを作成
			IDispatch xlSheet = (IDispatch)xlApp.get("ActiveSheet");

			Object[] arglist = new Object[1];
			arglist[0] = "A1";
			IDispatch xlRange = (IDispatch)xlSheet.get("Range",arglist);
			xlRange.put("Value","=1/0");

			Object value = xlRange.get("Value");
			System.out.println("value="+value);
			if (value != null) System.out.println("value.getClass()="+value.getClass());

			System.out.println("エラーの場合、VariantErrorを返すようにしました。");
			System.out.println("[Enter]を押してください。");
			System.in.read();

			Object[] arglist3 = new Object[3];
			arglist3[0] = new Boolean(false);
			arglist3[1] = null;
			arglist3[2] = new Boolean(false);
			xlBook.method("Close", arglist3);

			xlApp.method("Quit", null);

		}
		catch(Exception e) { e.printStackTrace(); }
		finally { rm.release();  }
	}
}
