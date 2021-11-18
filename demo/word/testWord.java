import jp.ne.so_net.ga2.no_ji.jcom.*;

/**
	���[�h�̃T���v��
	2001.07.04
*/
public class testWord {

	public static void main(String[] args) throws Exception {
		ReleaseManager rm = new ReleaseManager();
		try {
			System.out.println("Word���N����...");
			IDispatch wdApp = new IDispatch(rm, "Word.Application");  // EXCEL�{��
			wdApp.put("Visible", new Boolean(true));	// '�f�t�H���g��False(�\�����Ȃ�)

			IDispatch wdDocuments = (IDispatch)wdApp.get("Documents");
			Object[] arglist1 = new Object[1];
			String userdir = System.getProperty("user.dir");	// user.dir="E:\USERS\java\test"
			arglist1[0] = userdir+"\\jcom.doc";
			IDispatch wdDocument = (IDispatch)wdDocuments.method("Open", arglist1);
			String fullname = (String)wdDocument.get("FullName");
			System.out.println("fullname="+fullname);

			// �P������Ă݂�
			IDispatch wdWords = (IDispatch)wdDocument.get("Words");
			int word_count = ((Integer)wdWords.get("Count")).intValue();
			for(int i=0; i<word_count; i++) {
				Object[] index = new Object[1];
				index[0] = new Integer(i+1);		// COM�R���N�V�����͂P����n�܂�
				IDispatch wdWord = (IDispatch)wdWords.method("Item", index);
				String text = (String)wdWord.get("Text");
				System.out.println(text);
			}

			// �\�����Ă݂�
			IDispatch wdTables = (IDispatch)wdDocument.get("Tables");
			System.out.println(wdTables);
			int table_count = ((Integer)wdTables.get("Count")).intValue();
			System.out.println("�\�̐�="+table_count);
			for(int i=0; i<table_count; i++) {
				Object[] index = new Object[1];
				index[0] = new Integer(i+1);		// COM�R���N�V�����͂P����n�܂�
				IDispatch wdTable = (IDispatch)wdTables.method("Item", index);
				System.out.println(""+i+"="+wdTable);
				// �ꉞ�A�\�͎��邯��ǁE�E�E
			}

			// �v�����^�ɏo��
			//wdDocument.method("PrintOut", null);	// ���얢�m�F

			System.out.println("�R�b��ɏI�����܂��B");
			Thread.sleep(3000);	// 3sec
			wdApp.method("Quit", null);
			System.out.println("���Ò��A���肪�Ƃ��������܂����B");
		}
		catch(Exception e) { e.printStackTrace(); }
		finally { rm.release(); }
	}
}
