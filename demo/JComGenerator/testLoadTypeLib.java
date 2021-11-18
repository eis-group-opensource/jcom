import jp.ne.so_net.ga2.no_ji.jcom.*;

public class testLoadTypeLib {
	public static void main(String[] args) throws Exception {
		String file = "D:\\Office97\\Office\\EXCEL.EXE";
		if(args.length > 0) file = args[0];
		ReleaseManager rm = new ReleaseManager();
		try {
			ITypeLib typeLib = ITypeLib.loadTypeLib(rm, file);
			String[] docs = typeLib.getDocumentation(-1);
			System.out.println(docs[0]);
			System.out.println(docs[1]);
			System.out.println(docs[2]);
			System.out.println(docs[3]);

			System.out.println("Ç≤ê√íÆÅAÇ†ÇËÇ™Ç∆Ç§Ç≤Ç¥Ç¢Ç‹ÇµÇΩÅB");
		}
		catch(Exception e) { e.printStackTrace(); }
		finally { rm.release(); }
	}
}
