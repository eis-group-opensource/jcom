import jp.ne.so_net.ga2.no_ji.jcom.*;

public class testJCom {
	public static void main(String[] args) throws Exception {
		ReleaseManager rm = new ReleaseManager();
		try {
			System.out.println("EXCEL���N����...");
			IDispatch xlApp = new IDispatch(rm, "Excel.Application");  // EXCEL�{��
			xlApp.put("Visible", new Boolean(true));	// '�f�t�H���g��False(�\�����Ȃ�)
			IDispatch xlBooks = (IDispatch)xlApp.get("Workbooks");
			IDispatch xlBook = (IDispatch)xlBooks.method("Add",null);	// �V�����u�b�N���쐬
			IDispatch xlSheet = (IDispatch)xlApp.get("ActiveSheet");

			System.out.println("�Z��A1�ɕ�������Z�b�g");
			Object[] arglist = new Object[1];
			arglist[0] = "A1";
			IDispatch xlRange = (IDispatch)xlSheet.get("Range",arglist);
			xlRange.put("Value","JCom���������I(^o^)");

			System.out.println("[Enter]�������Ă��������B�I�����܂�");
			System.in.read();

			Object[] arglist3 = new Object[3];
			arglist3[0] = new Boolean(false);
			arglist3[1] = null;
			arglist3[2] = new Boolean(false);
			xlBook.method("Close", arglist3);
			xlApp.method("Quit", null);
			System.out.println("���Ò��A���肪�Ƃ��������܂����B");
		}
		catch(Exception e) { e.printStackTrace(); }
		finally { rm.release(); }
	}
}
