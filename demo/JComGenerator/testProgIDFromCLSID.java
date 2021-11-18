import jp.ne.so_net.ga2.no_ji.jcom.*;
import jp.ne.so_net.ga2.no_ji.jcom.excel8.*;

public class testProgIDFromCLSID {
	public static void main(String[] args) throws Exception {
		if(args.length!=1) {
			System.out.println("usage: testTypeLib <ProgID>");
			System.out.println("例  testTypeLib Excel.Application");
			return;
		}
		String progID = args[0];
		GUID CLSID = Com.getCLSIDFromProgID(progID);
		System.out.println("CLSID="+CLSID);
		String progid = Com.getProgIDFromCLSID(CLSID);
		System.out.println("ProgID="+progid);
	}
}
/*
>java -classpath %CLASSPATH%;../../jcom.jar testProgIDFromCLSID Excel.Application
CLSID={00024500-0000-0000-C000-000000000046}
ProgID=Excel.Application.8		←バージョン付きのProgIDを返す。
*/
