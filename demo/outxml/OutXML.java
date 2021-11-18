import jp.ne.so_net.ga2.no_ji.jcom.excel8.*;
import jp.ne.so_net.ga2.no_ji.jcom.*;
import java.io.*;

/*
	Excel�p���b�p���g�����AJCom�̃T���v���v���O����
	Excel����XML�ɕϊ�����B
*/
public class OutXML {

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

	// �w�肳�ꂽ�t�@�C�����猩�ς�����擾����
	public boolean getEstimate(String fname) {
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
			System.out.println("�Z������ݒ荀�ڂ��擾");
			company     = xlRange.Item(4,2).Text();
			section1    = xlRange.Item(5,2).Text();
			section2    = xlRange.Item(6,2).Text();
			custmer     = xlRange.Item(8,2).Text();
			validperiod = xlRange.Item(14,4).Text();
			createdate  = xlRange.Item(3,8).Text();
			estimatedNo = xlRange.Item(5,8).Text();
			charge      = xlRange.Item(7,8).Text();
			for(int i=0; i<15; i++) {
				itemname[i] = xlRange.Item(22+i,3).Text();
				itemtype[i] = xlRange.Item(22+i,4).Text();
				try {
					itemprice[i] = Integer.parseInt(xlRange.Item(22+i,5).Text());
					itemcount[i] = Integer.parseInt(xlRange.Item(22+i,6).Text());
				}
				catch(NumberFormatException e) {
					itemprice[i] = 0;
					itemcount[i] = 0;
				}
				itemmemo[i] = xlRange.Item(22+i,8).Text();
			}
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

	public static void main(String[] args) throws IOException {
		OutXML est = new OutXML();
		// XML�t�@�C�����쐬
		System.out.println("estimate.xml ���쐬");
		PrintWriter out = 
				new PrintWriter(new BufferedWriter(new FileWriter("estimate.xml")));

		// XML�w�b�_���o��	output XML header
		out.print("<?xml version=\"1.0\" standalone=\"no\"?>\n");
		out.print("<!DOCTYPE JCom:Estimate SYSTEM \"estimate.dtd\">\n");
		out.print("<!-- generated by OutXML.class -->\n");
		out.println();

		// �J�����g�f�B���N�g���� .xls �t�@�C���������������Ă���
		File path = new File("./");
		String[] filenames = path.list();
		for(int f=0; f<filenames.length; f++) {
			String fname = filenames[f];
			if(! fname.toLowerCase().endsWith(".xls")) continue;
			if( fname.equalsIgnoreCase("estimate.xls")) continue;
			System.out.println(fname);
			// Excel�͓Ǝ��̊������̂ŁA�t�@�C�������΃p�X�ɂ��ēn��
			boolean rc = est.getEstimate((new File(fname)).getCanonicalPath());
			if(rc) {
				out.print("<JCom:Estimate>\n");
				out.print("\t<JCom:Company>"+est.company+"</JCom:Company>\n");
				out.print("\t<JCom:Section1>"+est.section1+"</JCom:Section1>\n");
				out.print("\t<JCom:Section2>"+est.section2+"</JCom:Section2>\n");
				out.print("\t<JCom:Custmer>"+est.custmer+"</JCom:Custmer>\n");
				out.print("\t<JCom:ValidPeriod>"+est.validperiod+"</JCom:ValidPeriod>\n");
				out.print("\t<JCom:CreateDate>"+est.createdate+"</JCom:CreateDate>\n");
				out.print("\t<JCom:EstimatedNo>"+est.estimatedNo+"</JCom:EstimatedNo>\n");
				out.print("\t<JCom:Charge>"+est.charge+"</JCom:Charge>\n");
				loop:
				for(int i=0; i<15; i++) {
					if(est.itemcount[i] == 0) break loop;
					out.print("\t<JCom:Item>\n");
					out.print("\t\t<JCom:ItemName>"+est.itemname[i]+"</JCom:ItemName>\n");
					out.print("\t\t<JCom:ItemType>"+est.itemtype[i]+"</JCom:ItemType>\n");
					out.print("\t\t<JCom:ItemPrice>"+est.itemprice[i]+"</JCom:ItemPrice>\n");
					out.print("\t\t<JCom:ItemCount>"+est.itemcount[i]+"</JCom:ItemCount>\n");
					out.print("\t\t<JCom:ItemMemo>"+est.itemmemo[i]+"</JCom:ItemMemo>\n");
					out.print("\t</JCom:Item>\n");
				}
				out.print("</JCom:Estimate>\n");
				out.println();
			}
		}
		out.close();
		System.out.println("�������܂���");
	}
}