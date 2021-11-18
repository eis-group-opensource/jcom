import jp.ne.so_net.ga2.no_ji.jcom.*;

public class testSheetCopy {
	public static void main(String[] args) throws Exception {
		ReleaseManager rm = new ReleaseManager();
		try {
			System.out.println("EXCEL���N����...");
			IDispatch xlApp = new IDispatch(rm, "Excel.Application");
			xlApp.put("Visible", new Boolean(true));
			IDispatch xlBooks = (IDispatch)xlApp.get("Workbooks");
			IDispatch xlBook = (IDispatch)xlBooks.method("Add",null);	// create new book
			IDispatch xlSheet = (IDispatch)xlApp.get("ActiveSheet");

			// set string to cell A1 .
			System.out.println("�Z��A1�ɕ�������Z�b�g");
			Object[] arglist = new Object[1];
			arglist[0] = "A1";
			IDispatch xlRange = (IDispatch)xlSheet.get("Range",arglist);
			xlRange.put("Value","JCom (^o^)");

			// copy cell from A1 to B2 .
			// �Z�����R�s�[���Ă݂�B�P��Z��
			System.out.println("�Z��A1�̓��e��B1�ɃR�s�[");
			Object[] copyargs = new Object[1];
			copyargs[0] = (IDispatch)xlSheet.get("Range", new Object[] {"B2"});
			xlRange.method("Copy", copyargs);

			// copy cells from A1:B2 to C1:D2 .
			// �Z�����R�s�[���Ă݂�B�����Z�� A1:B2�� C1:D2�փR�s�[
			System.out.println("�Z��A1:B2�̓��e��C1:D2�փR�s�[");
			IDispatch xlRangeA1B2 = (IDispatch)xlSheet.get("Range",new Object[] {"A1:B2"});
			copyargs = new Object[1];
			copyargs[0] = (IDispatch)xlSheet.get("Range", new Object[] {"C1"});
			xlRangeA1B2.method("Copy", copyargs);

			// copy sheet.
			// �V�[�g���R�s�[���Ă݂�
			System.out.println("�V�[�g���R�s�[");
			copyargs = new Object[2];
			copyargs[0] = null;
			copyargs[1] = xlSheet;
			xlSheet.method("Copy", copyargs);

			System.out.println("Hit [Enter] key to exit.");
			System.in.read();

			// quit.
			// �I��������B
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
