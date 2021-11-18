import jp.ne.so_net.ga2.no_ji.jcom.*;
import jp.ne.so_net.ga2.no_ji.jcom.excel8.*;

public class testSheetCopy2 {
	public static void main(String[] args) throws Exception {
		ReleaseManager rm = new ReleaseManager();
		try {
			System.out.println("EXCEL���N����...");
			ExcelApplication xlApp = new ExcelApplication(rm);
			xlApp.Visible(true);
			ExcelWorkbooks xlBooks = xlApp.Workbooks();
			ExcelWorkbook xlBook = xlBooks.Add();	// create new book
			ExcelWorksheet xlSheet = xlApp.ActiveSheet();

			// set string to cell A1 .
			System.out.println("�Z��A1�ɕ�������Z�b�g");
			ExcelRange xlRange = xlSheet.Range("A1");
			xlRange.Value("JCom (^o^)");

			// copy cell from A1 to B2 .
			// �Z�����R�s�[���Ă݂�B�P��Z��
			System.out.println("�Z��A1�̓��e��B1�ɃR�s�[");
			xlRange.Copy(xlSheet.Range("B2"));

			// copy cells from A1:B2 to C1:D2 .
			// �Z�����R�s�[���Ă݂�B�����Z�� A1:B2�� C1:D2�փR�s�[
			System.out.println("�Z��A1:B2�̓��e��C1:D2�փR�s�[");
			ExcelRange xlRangeA1B2 = xlSheet.Range("A1:B2");
			xlRangeA1B2.Copy(xlSheet.Range("C1"));

			// copy sheet.
			// �V�[�g���R�s�[���Ă݂�
			System.out.println("�V�[�g���R�s�[");
			xlSheet.Copy(null, xlSheet);

			System.out.println("Hit [Enter] key to exit.");
			System.in.read();

			// quit.
			// �I��������B
			xlBook.Close(false, null, false);
			xlApp.Quit();
		}
		catch(Exception e) { e.printStackTrace(); }
		finally { rm.release(); }
	}
}
