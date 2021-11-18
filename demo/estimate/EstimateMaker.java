import jp.ne.so_net.ga2.no_ji.jcom.excel8.*;
import jp.ne.so_net.ga2.no_ji.jcom.*;
import java.io.File;
import java.util.Date;

/*
	Excel�p���b�p���g�����AJCom�̃T���v���v���O����
*/
public class EstimateMaker {

	public boolean print_enable = false;	// ������邩�ǂ���
	// �ݒ荀��
	public String company     = "";				// ��Ж�
	public String section1    = "";				// ����1
	public String section2    = "";				// ����2
	public String custmer     = "";				// �ڋq��
	public String validperiod = "";				// ���ϗL������
	public String createdate  = "";				// �쐬��
	public String estimatedNo = "";				// ���Ϗ�No.
	public String charge      = "";				// �S����
	public String[] itemname = new String[15];	// �i��x15��
	public String[] itemtype = new String[15];	// �^��x15��
	public int[] itemprice   = new int[15];		// �P��x15��
	public int[] itemcount   = new int[15];		// ����x15��
	public String[] itemmemo = new String[15];	// ���lx15��

	public boolean makeEstimate(String fname) {
		ReleaseManager rm = new ReleaseManager();
		try {
			
			System.out.println("EXCEL���N����...");
			// ���łɗ����オ���Ă���ƁA�V�����E�B���h�E�ŊJ���B
			ExcelApplication excel = new ExcelApplication(rm);
			excel.Visible(true);

			ExcelWorkbooks xlBooks = excel.Workbooks();
			ExcelWorkbook xlBook = xlBooks.Open(fname);
			ExcelWorksheet xlSheet = excel.ActiveSheet();
			ExcelRange xlRange = xlSheet.Cells();
			
			// �ݒ荀�ڂ��Z���ɑ��
			System.out.println("�ݒ荀�ڂ��Z���ɑ��");
			xlRange.Item(4,2).Value(company);
			xlRange.Item(5,2).Value(section1);
			xlRange.Item(6,2).Value(section2);
			xlRange.Item(8,2).Value(custmer);
			xlRange.Item(14,4).Value(validperiod);
			xlRange.Item(3,8).Value(createdate);
			xlRange.Item(5,8).Value(estimatedNo);
			xlRange.Item(7,8).Value(charge);
			for(int i=0; i<15; i++) {
				xlRange.Item(22+i,3).Value(itemname[i]);
				xlRange.Item(22+i,4).Value(itemtype[i]);
				if(itemcount[i]!=0) {
					xlRange.Item(22+i,5).Value(itemprice[i]);
					xlRange.Item(22+i,6).Value(itemcount[i]);
				}
				else {
					xlRange.Item(22+i,5).Value("");
					xlRange.Item(22+i,6).Value("");
				}
				xlRange.Item(22+i,8).Value(itemmemo[i]);
			}

			// �v�����^�ɏo�͂���ꍇ�̓R�����g���͂����Ă��������B
			// �f�t�H���g�̃v�����^�ɏo�͂���܂��B
			if(print_enable) {
				System.out.println("�v�����^�Ɉ�����܂��B");
				xlSheet.PrintOut();
			}

			// �t�@�C���ɕۑ�����ꍇ�̓R�����g���O���Ă��������B
			// �f�B���N�g�����w�肵�Ȃ��ꍇ�́A(My Documents)�ɕۑ�����܂��B
			System.out.println("�t�@�C���ɕۑ����܂��B");
			xlBook.Save();

			xlBook.Close(false,null,false);
			excel.Quit();
		}
		catch(Exception e) {
			e.printStackTrace();
			return false;	// ���s
		}
		finally { rm.release(); }
		return true;
	}

	public static void main(String[] args) {
		EstimateMaker est = new EstimateMaker();
		est.company     = "�܂����낻�ӂ�";			// ��Ж�
		est.section1    = "���Ղ�J������";			// ����1
		est.section2    = "�܂�����O���[�v";		// ����2 �������邾��I
		est.custmer     = "�т� ������ �l";			// �ڋq��
		est.validperiod = "2000/09/05";				// ���ϗL������
		est.createdate  = "2000/08/06";				// �쐬��
		est.estimatedNo = "PL1234-56-7890";			// ���Ϗ�No.
		est.charge      = "�n�� �`��";				// �S����
		for(int i=0; i<15; i++) {
			est.itemname[i]  = "-";		// �i��x15
			est.itemtype[i]  = "";		// �^��x15
			est.itemprice[i] = 0;		// �P��x15
			est.itemcount[i] = 0;		// ����x15
			est.itemmemo[i]  = "";		// ���lx15
		}
		// ����1
		est.itemname[0] = "�֐��c���[Ver14.40";
		est.itemtype[0] = "FT-1440S";
		est.itemprice[0] = 1500;
		est.itemcount[0] = 1;
		est.itemmemo[0]  = "�\�[�X�t";
		// ����2
		est.itemname[1] = "JCom 2.00";
		est.itemtype[1] = "JC-200";
		est.itemprice[1] = 800;
		est.itemcount[1] = 4;
		est.itemmemo[1]  = "";
		// ����3
		est.itemname[2] = "SYLBIS ����";
		est.itemtype[2] = "SY-1A";
		est.itemprice[2] = 3000;
		est.itemcount[2] = 1;
		est.itemmemo[2]  = "DirectX�Ή�";
		// �l��ݒ聕���.
		est.print_enable = false;
		// ��Ɨp�̃t�@�C�����쐬�B���Ϗ�No�Ɠ����ɂ���B
		try {
			String workfile = ".\\"+est.estimatedNo+".xls";
			System.out.println("�t�@�C�����R�s�[ "+workfile);
			FileCopy.copy(".\\estimate.xls", workfile);
			// ���Ϗ����쐬�BExcel�͓Ǝ��̊������̂ŁA�t�@�C�������΃p�X�ɂ��ēn��
			System.out.println("���Ϗ��쐬");
			boolean rc = est.makeEstimate((new File(workfile)).getCanonicalPath());
			if(rc)
				System.out.println("�������܂���");
			else
				System.out.println("���s���܂���(;_;)");
		}
		catch(Exception e) {
			e.printStackTrace();
			System.out.println("���s���܂���(T_T)");
		}
	}
}
