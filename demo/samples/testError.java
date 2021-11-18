import jp.ne.so_net.ga2.no_ji.jcom.*;

public class testError {
	public static void main(String[] args) throws Exception {
		ReleaseManager rm = new ReleaseManager();
		try {
			System.out.println("EXCEL���N����...");
			IDispatch xlApp = new IDispatch(rm, "Excel.Application");  // EXCEL�{��
			xlApp.put("Visible", new Boolean(true));	// '�f�t�H���g��False(�\�����Ȃ�)

			IDispatch xlBooks = (IDispatch)xlApp.get("Workbooks");
			
			IDispatch xlBook = (IDispatch)xlBooks.method("Add",null);	// �V�����u�b�N���쐬
			IDispatch xlSheet = (IDispatch)xlApp.get("ActiveSheet");

			Object[] arglist = new Object[1];
			arglist[0] = "A1";
			IDispatch xlRange = (IDispatch)xlSheet.get("Range",arglist);
			xlRange.put("Value","=1/0");

			Object value = xlRange.get("Value");
			System.out.println("value="+value);
			if (value != null) System.out.println("value.getClass()="+value.getClass());

			System.out.println("�G���[�̏ꍇ�AVariantError��Ԃ��悤�ɂ��܂����B");
			System.out.println("[Enter]�������Ă��������B");
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
