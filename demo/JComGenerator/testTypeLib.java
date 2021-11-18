import jp.ne.so_net.ga2.no_ji.jcom.*;

public class testTypeLib {
	
	public static String getProgID(IUnknown unknown) {
		try {
			IPersist persist = (IPersist)unknown.queryInterface(IPersist.class, IPersist.IID);
			if(persist==null) return null;
			GUID clsid = persist.getClassID();
			return Com.getProgIDFromCLSID(clsid);
		}
		catch(JComException e) { e.printStackTrace(); }
		return null;
	}
	/**
		指定したProgIDのタイプライブラリを見る
	*/
	public static void main(String[] args) throws Exception {
		if(args.length!=1) {
			System.out.println("usage: testTypeLib <ProgID>");
			System.out.println("例  testTypeLib Excel.Application");
			return;
		}
		String progID = args[0];
		ReleaseManager rm = new ReleaseManager();
		try {
			IDispatch disp = new IDispatch(rm, progID);
			ITypeInfo typeinfo = disp.getTypeInfo();
			ITypeLib typelib = typeinfo.getTypeLib();
			// ドキュメントを表示
			String[] docs = typelib.getDocumentation(-1);
			System.out.println("docs[0]="+docs[0]);
			System.out.println("docs[1]="+docs[1]);
			System.out.println("docs[2]="+docs[2]);
			System.out.println("docs[3]="+docs[3]);
			// TLIBATTRを表示
			ITypeLib.TLibAttr attr = typelib.getTLibAttr();
			System.out.println("TLIBATTR="+attr);
			// ITypeInfoの数
			int infocount = typelib.getTypeInfoCount();
			System.out.println("ITypeInfoの数="+infocount);
			if(true) {
				for(int i=0; i<infocount; i++) {
					ITypeInfo info = typelib.getTypeInfo(i);
					docs = info.getDocumentation(-1);
//					System.out.print("ProgID="+getProgID(info));
					System.out.println(docs[0]);
				}
			}
		}
		catch(Exception e) { e.printStackTrace(); }
		finally { rm.release(); }
	}
}
