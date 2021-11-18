import jp.ne.so_net.ga2.no_ji.jcom.*;
import jp.ne.so_net.ga2.no_ji.jcom.excel8.*;

class testEnum {

	public static void main(String[] args) throws Exception {
		ReleaseManager rm = new ReleaseManager();
		try {
			System.out.println("EXCEL���N����...");
			// ���łɗ����オ���Ă���ƁA�V�����E�B���h�E�ŊJ���B
			ExcelApplication excel = new ExcelApplication(rm);
			excel.Visible(true);

			ExcelWorkbooks xlBooks = excel.Workbooks();
			ExcelWorkbook xlBook = xlBooks.Add();	// �V�����u�b�N���쐬
			ExcelWorksheets xlSheets = xlBook.Worksheets();

			// �V�[�g����IEnumVARIANT�𐶐��BEnum���ۂ�����
			System.out.println("���ׂẴV�[�g����񋓂��Ă݂�");
			IEnumVARIANT enum = xlSheets.NewEnum();
			for(;;) {
				IDispatch disp = (IDispatch)enum.next();
				if(disp==null) break;
				ExcelWorksheet xlSheet = new ExcelWorksheet(disp);
				System.out.println(""+xlSheet.Name());
			}

			System.out.println("�ʂ̕��@�œ������Ƃ�����Ă݂�");
			enum.reset();	// �ŏ�����
			Object[] ary = enum.next(100);	// �ő�P�O�O�̃I�u�W�F�N�g���擾
			for(int i=0; i<ary.length; i++) {
				ExcelWorksheet xlSheet = new ExcelWorksheet((IDispatch)ary[i]);
				System.out.println(""+xlSheet.Name());
			}

			System.out.println("[Enter]�������Ă��������B�I�����܂�");
			System.in.read();

			xlBook.Close(false,null,false);
			excel.Quit();

			System.out.println("���Ò��A���肪�Ƃ��������܂����B");
		}
		catch(Exception e) { e.printStackTrace(); }
		finally { rm.release(); }
	}
}
