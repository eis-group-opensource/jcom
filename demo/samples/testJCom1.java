import jp.ne.so_net.ga2.no_ji.jcom.*;

public class testJCom1 {
	public static void main(String[] args) throws Exception {
		ReleaseManager rm = new ReleaseManager();
		try {
			System.out.println("EXCEL��\�������ɋN��...");
			IDispatch xlApp = new IDispatch(rm, "Excel.Application");  // EXCEL�{��

			System.out.println("�\������Ȃ��̂́AVisible��false�̂܂܂�����ł��B");
			System.out.println("�^�X�N�}�l�[�W�����łd�����������N�����Ă���̂��m�F�ł��܂��B");
			System.out.println("\n[Enter]�������Ă��������B�I�����܂�");
			System.in.read();
			xlApp.invoke("Quit", null);

			xlApp.release();
			System.out.println("���Ò��A���肪�Ƃ��������܂����B");
		}
		catch(Exception e) { e.printStackTrace(); }
		finally { rm.release(); }
	}
}
