import jp.ne.so_net.ga2.no_ji.jcom.*;
import java.io.*;
import java.util.*;
import java.text.DateFormat;

public class ListTypes {

	static String tab = "\t";
	void listTypes(ITypeLib typeLib) {
		try {
			// ファイルをオープン(CSVで保存しているところが、心憎い演出)
			PrintWriter out = 
					new PrintWriter(new BufferedWriter(new FileWriter("ExcelTypes.csv")));

			int infocount = typeLib.getTypeInfoCount();
			for(int i=0; i<infocount; i++) {
				ITypeInfo info = typeLib.getTypeInfo(i);
				out.print(info.getDocumentation(-1)[0]+tab);
				ITypeInfo.TypeAttr attr = info.getTypeAttr();
				out.print(attr.getIID()+tab);
				switch(attr.getTypeKind()) {
					case ITypeInfo.TypeAttr.TKIND_ENUM:			out.print("TKIND_ENUM");		break;
					case ITypeInfo.TypeAttr.TKIND_RECORD:		out.print("TKIND_RECORD");		break;
					case ITypeInfo.TypeAttr.TKIND_MODULE:		out.print("TKIND_MODULE");		break;
					case ITypeInfo.TypeAttr.TKIND_INTERFACE:	out.print("TKIND_INTERFACE");	break;
					case ITypeInfo.TypeAttr.TKIND_DISPATCH:		out.print("TKIND_DISPATCH");	break;
					case ITypeInfo.TypeAttr.TKIND_COCLASS:		out.print("TKIND_COCLASS");		break;
					case ITypeInfo.TypeAttr.TKIND_ALIAS:		out.print("TKIND_ALIAS");		break;
					case ITypeInfo.TypeAttr.TKIND_UNION:		out.print("TKIND_UNION");		break;
				}
				out.println(tab+"Func="+attr.getFuncs()+tab+"Var="+attr.getVars());
			}
			out.close();
		}
		catch(Exception e) { e.printStackTrace(); }
	}

	public static void main(String[] args) throws Exception {
		ReleaseManager rm = new ReleaseManager();
		try {
			System.out.println("EXCELを起動中...");
			IDispatch xlApp = new IDispatch(rm, "Excel.Application");  // EXCEL本体
			xlApp.put("Visible", new Boolean(true));	// 'デフォルトはFalse(表示しない)

			ITypeInfo typeInfo = xlApp.getTypeInfo();
			ITypeLib typeLib = typeInfo.getTypeLib();
			ListTypes listTypes = new ListTypes();
			listTypes.listTypes(typeLib);
			xlApp.invoke("Quit", null);
			System.out.println("ご静聴、ありがとうございました。");
		}
		catch(Exception e) { e.printStackTrace(); }
		finally { rm.release(); }
	}
}
