import jp.ne.so_net.ga2.no_ji.jcom.excel8.*;
import jp.ne.so_net.ga2.no_ji.jcom.*;
import java.io.File;
import java.util.Date;

/* Excel�p���b�p���g�����AJCom�̃T���v���v���O���� 
	�w�i�F�A�{�[�_�[�̃e�X�g
*/
class testInterior {
	public static void main(String[] args) throws Exception {
		ReleaseManager rm = new ReleaseManager();
		try {
			System.out.println("EXCEL���N����...");
			// ���łɗ����オ���Ă���ƁA�V�����E�B���h�E�ŊJ���B
			ExcelApplication excel = new ExcelApplication(rm);
			excel.Visible(true);
			// �F��ȏ���\��
			System.out.println("Version="+excel.Version());
			System.out.println("UserName="+excel.UserName());
			System.out.println("Caption="+excel.Caption());
			System.out.println("Value="+excel.Value());

			ExcelWorkbooks xlBooks = excel.Workbooks();
			ExcelWorkbook xlBook = xlBooks.Add();	// �V�����u�b�N���쐬
			
			// ���ׂẴt�@�C����񋓂��Ă݂�
			System.out.println("���݂̃f�B���N�g���̃t�@�C�����Z���ɐݒ�");
			ExcelWorksheets xlSheets = xlBook.Worksheets();
			ExcelWorksheet xlSheet = xlSheets.Item(1);
			ExcelRange xlRange = xlSheet.Cells();

			xlRange.Item(1,1).Value("�t�@�C����" );
			xlRange.Item(1,2).Value("�T�C�Y" );
			xlRange.Item(1,3).Value("�ŏI�X�V����");
			xlRange.Item(1,4).Value("�f�B���N�g��");
			xlRange.Item(1,5).Value("�t�@�C��");
			xlRange.Item(1,6).Value("�ǂݍ��݉�");
			xlRange.Item(1,7).Value("�������݉�");
//			xlRange.Item(1,8).Value("�B���t�@�C��");

			// �჌�x���C���^�[�t�F�[�X���g���āA�T�|�[�g����Ă��Ȃ��I�u�W�F�N�g�ɃA�N�Z�X���Ă݂�B
			// �Z���̔w�i�F�ɐF��ݒ�B
			IDispatch interior = (IDispatch)xlRange.Item(1,1).get("Interior");
			interior.put("Color", new Integer(0xFFFF00));	// GGBBRR cyan
			// �r����ݒ�B
			IDispatch borders = (IDispatch)xlRange.Item(1,1).get("Borders");
			Object[] border_args = new Integer[1];
			border_args[0] = new Integer(9);
			IDispatch border = (IDispatch)borders.get("Item",border_args);	// XlBordersIndex.xlEdgeBottom = 9
			border.put("LineStyle", new Integer(1));	// XlLineStyle.xlContinuous = 1


			File path = new File("./");
			String[] filenames = path.list();
			for(int i=0; i<filenames.length; i++) {
				File file = new File(filenames[i]);
				System.out.println(file);
				xlRange.Item(i+2,1).Value( file.getName() );				// �t�@�C�����p�X����
				xlRange.Item(i+2,2).Value( (int)file.length() );			// �t�@�C���T�C�Y
				xlRange.Item(i+2,3).Value( new Date(file.lastModified()) );	// �ŏI�X�V����
				xlRange.Item(i+2,4).Value( file.isDirectory()?"Yes":"No" );	// �f�B���N�g�����H
				xlRange.Item(i+2,5).Value( file.isFile()?"Yes":"No" );		// �t�@�C�����H
				xlRange.Item(i+2,6).Value( file.canRead()?"Yes":"No" );		// �ǂݎ����H
				xlRange.Item(i+2,7).Value( file.canWrite()?"Yes":"No" );	// �������݉��H
//				xlRange.Item(i+2,8).Value( file.isHidden()?"Yes":"No" );	// �B���t�@�C�����H (jdk1.2�ȍ~)
			}
			String expression = "=Sum(B2:B"+(filenames.length+1)+")";
			System.out.println("�����𖄂ߍ��݁A�t�@�C���T�C�Y�̍��v�����߂� "+expression);
			xlRange.Item(filenames.length+2,1).Value("���v");
			xlRange.Item(filenames.length+2,2).Formula(expression);
			xlRange.Columns().AutoFit();	// �������t�B�b�g������

			// �v�����^�ɏo�͂���ꍇ�̓R�����g���͂����Ă��������B
			// �f�t�H���g�̃v�����^�ɏo�͂���܂��B
//			System.out.println("�v�����^�Ɉ�����܂��B");
//			xlSheet.PrintOut();

			// �t�@�C���ɕۑ�����ꍇ�̓R�����g���O���Ă��������B
			// �f�B���N�g�����w�肵�Ȃ��ꍇ�́A(My Documents)�ɕۑ�����܂��B
//			System.out.println("�t�@�C���ɕۑ����܂��B(My Documents)\\testExcel.xls");
//			xlBook.SaveAs("testExcel.xls");

			System.out.println("10�b��ɏI�����܂�");
			Thread.sleep(10*1000);

			xlBook.Close(false,null,false);
			excel.Quit();

			System.out.println("���Ò��A���肪�Ƃ��������܂����B");
		}
		catch(Exception e) { e.printStackTrace(); }
		finally { rm.release(); }
	}
}
