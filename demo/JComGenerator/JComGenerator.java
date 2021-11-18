import jp.ne.so_net.ga2.no_ji.jcom.*;
import java.io.*;
import java.util.*;
import java.text.DateFormat;

/**
	指定したProgIDのJava-COMブリッジを生成するクラス。

	COMとJavaの概念的な対応
	ITypeLib	package
	ITypeInfo	class
	メソッド	メソッド

*/
public class JComGenerator {
	// 生成前に設定するもの
	protected String outdirs = null;		// 出力ディレクトリ"..\.."
	protected String packagePath = null;	// パッケージのパス"jp.ne.so_net.ga2.no_ji.jcom"
	protected String progID = null;			// ProgID
	protected ITypeInfo entryInfo = null;	// ProgIDで指定された型情報
	// 生成中に変わるモノ
	protected String currentPackageName = null;	// 現在のパッケージ名"Excel"
	protected ITypeLib currentLib = null;		// 現在のライブラリ
	protected ITypeInfo currentInfo = null;		// 現在の型情報

	public void setOutDirs(String outdirs) { this.outdirs = outdirs; }
	public void setPackagePath(String packagePath) { this.packagePath = packagePath; }

	/**
		パッケージ名 "jp.ne.so_net.ga2.no_ji.jcom.Excel"と出力ディレクトリ"..\.."から
		..\..\jp\ne\so_net\ga2\no_ji\jcom\Excel"を返す。
	*/
	public String getPath(String packagename) {
		String path = outdirs;
		if(packagename != null) {
			if(path.length() > 0 && ! path.endsWith(File.separator))
				path += File.separator;
			path += packagename.replace('.', File.separatorChar);	// '\\'
		}
		return path;
	}
	/**
		ディレクトリを作る。複数ディレクトリ対応
	*/
	public boolean makedirs(String packagename) {
		return (new File(getPath(packagename))).mkdirs();
	}
	/**
		ファイル名を返す。
		パッケージ名 "jp.ne.so_net.ga2.no_ji.jcom.Excel"と出力ディレクトリ"..\.."と
		クラス名"Application"から
		..\..\jp\ne\so_net\ga2\no_ji\jcom\Excel\Application.java"を返す。
	*/
	public String getFilename(String packagename, String classname) {
		String path = getPath(packagename);
		return path + File.separator + classname + ".java";
	}

	/**
		TYPEATTR.typekind により、場合分け
	*/
	public boolean generate(ITypeInfo info) {
		currentInfo = info;
		try {
			ITypeInfo.TypeAttr attr = info.getTypeAttr();
			// すべてはサポートしていない。(;_;)
			switch(attr.getTypeKind()) {
				case ITypeInfo.TypeAttr.TKIND_ENUM:			return generateEnum(info);
				case ITypeInfo.TypeAttr.TKIND_DISPATCH:		return generateDispatch(info);
				case ITypeInfo.TypeAttr.TKIND_COCLASS:		return generateCoClass(info);
				case ITypeInfo.TypeAttr.TKIND_RECORD:		
				case ITypeInfo.TypeAttr.TKIND_MODULE:		
				case ITypeInfo.TypeAttr.TKIND_INTERFACE:	// 3
				case ITypeInfo.TypeAttr.TKIND_ALIAS:		
				case ITypeInfo.TypeAttr.TKIND_UNION:		
//					return generateNoSupport(info);
					System.out.print(" @@@ 未対応 TypeKind="+attr.getTypeKind());
					break;
				default: 
					System.out.print(" @@@@うそ？！ "+attr.getTypeKind());
			}
			return true;
		}catch(Exception e) { e.printStackTrace(); }
		return false;
	}
	/**
		ファイルヘッダを出力
	*/
	void outHeader(PrintWriter out, String packagename) {
		// 現在日時を取得。
		Date now = Calendar.getInstance().getTime();
		String nowstr = DateFormat.getDateTimeInstance().format(now);
		out.println("// generated by "+this.getClass().toString()+" at "+nowstr);
		out.println("// JCom 2.20   URL=http://www02.so-net.ne.jp/~no-ji");
		if(packagename!=null) out.println("package "+packagename+";");
		out.println();
		out.println("import jp.ne.so_net.ga2.no_ji.jcom.*;");
		out.println();
	}

	/**
		サポートされていないタイプのインターフェースを生成します。
		TYPEATTR.typekind が TKIND_???? のモノに対して生成
		デバッグ用に出力します。
	*/
	public boolean generateNoSupport(ITypeInfo info) {
		String packagename = currentPackageName;
		try {
			String[] docs = info.getDocumentation(-1);
			// ファイルをオープン
			makedirs(packagename);
			String fname = getFilename(packagename, docs[0]);
			PrintWriter out = 
					new PrintWriter(new BufferedWriter(new FileWriter(fname)));
			// ファイルヘッダを出力
			outHeader(out, packagename);
			// 出力
			ITypeInfo.TypeAttr attr = info.getTypeAttr();
			out.println("@@@ 未対応 TypeKind="+attr.getTypeKind());
			switch(attr.getTypeKind()) {
				case ITypeInfo.TypeAttr.TKIND_ENUM:			out.print("TKIND_ENUM");		break;
				case ITypeInfo.TypeAttr.TKIND_DISPATCH:		out.print("TKIND_DISPATCH");	break;
				case ITypeInfo.TypeAttr.TKIND_RECORD:		out.print("TKIND_RECORD");		break;
				case ITypeInfo.TypeAttr.TKIND_MODULE:		out.print("TKIND_MODULE");		break;
				case ITypeInfo.TypeAttr.TKIND_INTERFACE:	out.print("TKIND_INTERFACE");	break;
				case ITypeInfo.TypeAttr.TKIND_COCLASS:		out.print("TKIND_COCLASS");		break;
				case ITypeInfo.TypeAttr.TKIND_ALIAS:		out.print("TKIND_ALIAS");		break;
				case ITypeInfo.TypeAttr.TKIND_UNION:		out.print("TKIND_UNION");		break;
			}
			// ファイルを閉じる
			out.close();
			return true;
		}catch(Exception e) { e.printStackTrace(); }
		return false;
	}

	/**
		列挙型のクラスを生成します。
		TYPEATTR.typekind が TKIND_ENUM のモノに対して生成
		varkind が VAR_CONST のものしか対応していませんが、
		Excelの場合、実質それしかないようです。
		@param	outdir		出力パス名(\なし)
	*/
	public boolean generateEnum(ITypeInfo info) {
		String packagename = currentPackageName;
		try {
			String[] docs = info.getDocumentation(-1);
			// ファイルをオープン
			makedirs(packagename);
			String fname = getFilename(packagename, docs[0]);
			PrintWriter out = 
					new PrintWriter(new BufferedWriter(new FileWriter(fname)));
			// ファイルヘッダを出力
			outHeader(out, packagename);
			// 関数ヘッダを出力
			out.println("/**");
			out.println("\tEnum "+docs[0]+((docs[1]!=null)?docs[1]:""));
			out.println("\t"+docs[3]+" "+docs[2]);
			out.println("*/");
			out.println("public class "+docs[0]+" {");
			// 標準のメンバを出力
			ITypeInfo.TypeAttr attr = info.getTypeAttr();
			// 関数を出力
			if(attr.getFuncs()!=0) {
				out.print("Enumには関数はないんじゃないの？");
				return false;
			}
			for(int f=0; f<attr.getVars(); f++) {
				ITypeInfo.VarDesc   vardesc   = info.getVarDesc(f);
				ITypeInfo.ElemDesc  elemVar   = vardesc.getVar();
				String[]            names     = vardesc.getNames();
				Object              value     = vardesc.getValue();
				switch(vardesc.getVarKind()) {
					case ITypeInfo.VarDesc.VAR_CONST:		// 2
						out.println("\tpublic static final "+getJavaType(elemVar.toString())+" "+ names[0]+" = "+value+";");
						break;
					case ITypeInfo.VarDesc.VAR_PERINSTANCE:	// 0
					case ITypeInfo.VarDesc.VAR_STATIC:		// 1
					case ITypeInfo.VarDesc.VAR_DISPATCH:	// 3
						out.print("\t/* @@@ 未対応 Enum "+vardesc.toString()+" */");
					break;
				}
			}
			out.println("}");
			// ファイルを閉じる
			out.close();
			return true;
		}catch(Exception e) { e.printStackTrace(); }
		return false;
	}

	//	TYPEATTR.typekind が TKIND_DISPATCH のモノに対して生成
	public boolean generateDispatch(ITypeInfo info) {
		String packagename = currentPackageName;
		// ファイル名は 
		try {
			String[] docs = info.getDocumentation(-1);
			// ファイルをオープン
			makedirs(packagename);
			String fname = getFilename(packagename, docs[0]);
			PrintWriter out = 
					new PrintWriter(new BufferedWriter(new FileWriter(fname)));
			// ファイルヘッダを出力
			outHeader(out, packagename);
			// クラスヘッダを出力
			out.println("/**");
			out.println("\tDispatch "+docs[0]+((docs[1]!=null)?docs[1]:""));
			out.println("\t"+docs[3]+" "+docs[2]);
			String formalProgID = null;
			if(info.equals(entryInfo)) {	// ProgID、CLSIDを表示
				GUID CLSID = Com.getCLSIDFromProgID(progID);
				formalProgID = Com.getProgIDFromCLSID(CLSID);
				out.println("\tProgID="+formalProgID);
				out.println("\tCLSID="+CLSID);
			}
			out.println("*/");
			out.println("public class "+docs[0]+" extends IDispatch {");
			// 標準のメンバを出力。interface ID
			ITypeInfo.TypeAttr attr = info.getTypeAttr();
			out.println("\t/**");
			out.println("\t\tインターフェースＩＤです。");
			out.println("\t\t@see\tGUID");
			out.println("\t*/");
			out.println("\tpublic static GUID IID = GUID.parse(\""+attr.getIID().toString()+"\");");
			out.println();
			// ProgIDからのエントリが可能なとき、
			if(info.equals(entryInfo)) {
				out.println("\t/**");
				out.println("\t\tProgIDから生成します。");
				out.println("\t*/");
				out.println("\tpublic "+docs[0]+"(ReleaseManager rm) throws JComException {");
				out.println("\t\tsuper(rm, \""+formalProgID+"\");");
				out.println("\t}");
				out.println();
			}
			// 標準のメンバを出力。IDispatchからの生成
			out.println("\t/**");
			out.println("\t\tIDispatchから生成。戻り値から生成するときに使用");
			out.println("\t*/");
			out.println("\tpublic "+docs[0]+"(IDispatch disp) { super(disp); }");
			out.println();
			
			// 関数を出力
			nextroop:
			for(int f=0; f<attr.getFuncs(); f++) {
				ITypeInfo.FuncDesc   funcdesc   = info.getFuncDesc(f);
				ITypeInfo.ElemDesc   elemFunc   = funcdesc.getFunc();
				ITypeInfo.ElemDesc[] elemParams = funcdesc.getParams();
				String[]             names      = funcdesc.getNames();
				// IDispatch以外の関数を生成する。QueryInterface()などはいらない
				if(isIDispatchFunction(names[0])) continue nextroop;
				// 関数ヘッダを生成。なるべく、多くの情報を出力してやる。
				out.print("\t/**\n");
				out.print("\t\t"+funcdesc.toString()+"\n");
				String[] funcdocs = info.getDocumentation(funcdesc.getMemID());
				if(funcdocs != null) {
					out.print("\t\t");
					if(funcdocs[0]!=null) out.print(funcdocs[0]+" ");
					if(funcdocs[1]!=null) out.print(funcdocs[1]+" ");
					if(funcdocs[2]!=null) out.print(funcdocs[2]+" ");
					if(funcdocs[3]!=null) out.print(funcdocs[3]+" ");
					out.print("\n");
				}
				out.print("\t*/\n");
				String flagname = "";
				switch(funcdesc.getInvokeKind()) {
					case ITypeInfo.FuncDesc.INVOKE_FUNC:
						flagname = "IDispatch.METHOD";
						break;
					case ITypeInfo.FuncDesc.INVOKE_PROPERTYGET:
						flagname = "IDispatch.PROPERTYGET";
						break;
					case ITypeInfo.FuncDesc.INVOKE_PROPERTYPUT:
						flagname = "IDispatch.PROPERTYPUT";
						break;
					case ITypeInfo.FuncDesc.INVOKE_PROPERTYPUTREF:
						out.print("\t// @@@ 未対応 INVOKE_PROPERTYPUTREF\n");
						continue nextroop;		// 次の処理へ
//							flagname = "IDispatch.PROPERTYPUTREF";
//							break;
				}
				// このライブラリ以外のコンポーネントを使っていないか？
				if(containStranger(funcdesc)) {
					out.print("\t// @@@ 未対応 外部のライブラリのコンポーネントを含んでいます\n");
					continue nextroop;
				}
				// メソッド宣言部
				out.print("\tpublic "+getJavaType(elemFunc.getTypeDesc())+" "+names[0]+"(");
				if(elemParams != null) {
					for(int p=0; p<elemParams.length; p++) {
						String argname = (names.length>p+1) ? names[p+1] : ("_"+p);		// 名前がないときは "_0"に
						out.print(getJavaType(elemParams[p].getTypeDesc())+" "+argname);
						if(p==elemParams.length-1)
							out.print(") throws JComException {\n");
						else {
							out.print(", ");
							// ５つごとに改行をいれてやる。小さな心遣い。
							if(p%5==4) out.print("\n\t\t\t\t\t");
						}
						// その型の身元を洗う
						checkUserDefined(info, elemParams[p].getTypeDesc());
					}
				}
				else {
					out.println(") throws JComException {");
				}
				// 引数を変換、呼び出す
				if(elemParams!=null && elemParams.length>0) {
					out.print("\t\tObject[] params = new Object["+(elemParams.length)+"];\n");
					for(int i=0; i<elemParams.length; i++) {
						out.print("\t\tparams["+i+"] = ");
						String argname = (names.length>i+1) ? names[i+1] : ("_"+i);		// 名前がないときは "_1"に
						out.print(generateArgument(elemParams[i], argname)+";\n");
					}
					out.print("\t\tObject rc = invoke(\""+names[0]+"\", "+flagname+", params);\n");
				}
				else {
					out.print("\t\tObject rc = invoke(\""+names[0]+"\", "+flagname+", null);\n");
				}
				// return ステートメントを出力
				out.print("\t\t"+generateReturnStatement(elemFunc,"rc")+"\n");
				out.println("\t}");
				// 戻り値の型の身元を洗う
				checkUserDefined(info, elemFunc.getTypeDesc());
			}
			out.println("}");
			// ファイルを閉じる
			out.close();
			return true;
		}catch(Exception e) { e.printStackTrace(); }
		return false;
	}

	//	TYPEATTR.typekind が TKIND_COCLASS のモノに対して生成
	public boolean generateCoClass(ITypeInfo info) {
		String packagename = currentPackageName;
		// ファイル名は 
		try {
			String[] docs = info.getDocumentation(-1);
			// ファイルをオープン
			makedirs(packagename);
			String fname = getFilename(packagename, docs[0]);
			PrintWriter out = 
					new PrintWriter(new BufferedWriter(new FileWriter(fname)));
			// ファイルヘッダを出力
			outHeader(out, packagename);
			// 関数ヘッダを出力
			out.println("/**");
			out.println("\tCoClass "+docs[0]+((docs[1]!=null)?docs[1]:""));
			out.println("\t"+docs[3]+" "+docs[2]);
			ITypeInfo.TypeAttr attr = info.getTypeAttr();
			for(int i=0; i<attr.getImplTypes(); i++) {
				ITypeInfo imp = info.getImplType(i);
				out.println("\t"+imp.getDocumentation(-1)[0]);
			}
			out.println("*/");
			ITypeInfo impl = info.getImplType(0);
			out.println("public class "+docs[0]+" extends "+impl.getDocumentation(-1)[0]+" {");
			// 標準のメンバを出力
			out.println("\t/**");
			out.println("\t\tIDispatchから生成。戻り値から生成するときに使用");
			out.println("\t*/");
			out.println("\tpublic "+docs[0]+"(IDispatch disp) { super(disp); }");
			out.println();
			out.println("}");
			// ファイルを閉じる
			out.close();
			return true;
		}catch(Exception e) { e.printStackTrace(); }
		return false;
	}
	/**
		IDispatchインターフェースで定義されている関数かどうかを調べる。
		単純に関数名で判断しているだけ。
	*/
	protected boolean isIDispatchFunction(String funcname) {
		String[] fnames = {
			"QueryInterface",
			"AddRef",
			"Release",
			"GetTypeInfoCount",
			"GetTypeInfo",
			"GetIDsOfNames",
			"Invoke" };
		for(int i=0; i<fnames.length; i++) {
			if(funcname.equals(fnames[i])) return true;
		}
		return false;
	}
	/**
		VARIANTの型からJavaの型に変換
	*/
	protected String getJavaType(String type) {
		if(type.equals("VT_INT")) return "int";				// 推奨されない
		if(type.equals("VT_UI1")) return "byte";
		if(type.equals("VT_I2")) return "short";
		if(type.equals("VT_I4")) return "int";
		if(type.equals("VT_R4")) return "float";
		if(type.equals("VT_R8")) return "double";
		if(type.equals("VT_BSTR")) return "String";
		if(type.equals("VT_VOID")) return "void";			// 推奨されない
		if(type.equals("VT_VARIANT")) return "Object";		// 推奨されない
		if(type.equals("VT_BOOL")) return "boolean";
		if(type.equals("VT_DISPATCH")) return "IDispatch";
		if(type.equals("VT_UNKNOWN")) return "IUnknown";
		if(type.equals("VT_DATE")) return "java.util.Date";
		if(type.equals("VT_PTR+VT_I4")) return "int[]";
		if(type.equals("VT_PTR+VT_R8")) return "double[]";
		if(type.equals("VT_PTR+VT_BOOL")) return "boolean[]";
		if(type.equals("VT_PTR+VT_BSTR")) return "String[]";

		if(type.startsWith("VT_PTR+VT_USERDEFINED")) {
			String clsname = type.substring(type.indexOf(':')+1);
			// 同じライブラリ内か？
			try {
				int hRefType = ITypeInfo.getRefTypeFromTypeDesc(type);
				ITypeLib reflib = currentInfo.getRefTypeInfo(hRefType).getTypeLib();
				if(reflib.getTLibAttr().equals(currentLib.getTLibAttr())) {
					return clsname;
				}
				// 違ったら・・・、フルパスで表示
				String packagename = getPackageName(packagePath, reflib);
				return packagename+"."+clsname;
			} catch(Exception e) { e.printStackTrace(); }
			return "@@@"+clsname;
		}

		// VT_PTRのないVT_USERDEFINEDをEnumと見なし、Enumをintとみなす。ちょっと乱暴かも？
		if(type.startsWith("VT_USERDEFINED")) return "int";
		return "@@@"+type;
	}
	/**
		他のライブラリのコンポーネントを含むか？
	*/
	protected boolean containStranger(ITypeInfo.FuncDesc funcdesc) {
		ITypeInfo.ElemDesc   elemFunc   = funcdesc.getFunc();
		if(isStranger(elemFunc.getTypeDesc())) return true;
		ITypeInfo.ElemDesc[] elemParams = funcdesc.getParams();
		if(elemParams == null) return false;
		for(int p=0; p<elemParams.length; p++) {
			if(isStranger(elemParams[p].getTypeDesc())) return true;
		}
		return false;
	}
	/**
		他のライブラリのコンポーネントか？
	*/
	protected boolean isStranger(String type) {
		if(! type.startsWith("VT_PTR+VT_USERDEFINED")) return false;
		String clsname = type.substring(type.indexOf(':')+1);
		// 同じライブラリ内か？
		try {
			int hRefType = ITypeInfo.getRefTypeFromTypeDesc(type);
			ITypeLib reflib = currentInfo.getRefTypeInfo(hRefType).getTypeLib();
			if(reflib.getTLibAttr().equals(currentLib.getTLibAttr())) return false;
		} catch(Exception e) { e.printStackTrace(); }
		// 違ったら・・・
		return true;
	}
	/**
		引数を設定する箇所を生成する
	*/
	String generateArgument(ITypeInfo.ElemDesc elem, String name) {
		String jtype = getJavaType(elem.getTypeDesc());
		if(jtype.equals("byte")) return "new Byte("+name+")";
		if(jtype.equals("short")) return "new Short("+name+")";
		if(jtype.equals("int")) return "new Integer("+name+")";
		if(jtype.equals("float")) return "new Float("+name+")";
		if(jtype.equals("double")) return "new Double("+name+")";
		if(jtype.equals("boolean")) return "new Boolean("+name+")";
		if(jtype.equals("String")) return name;
		if(jtype.equals("Object")) return name;			// 推奨されない
		if(jtype.equals("IUnknown")) return name;	// よいか？
		if(jtype.equals("IDispatch")) return name;	// よいか？
		if(elem.getTypeDesc().startsWith("VT_PTR+VT_USERDEFINED")) return name;
		if(jtype.equals("int[]")) return name;
		if(jtype.equals("double[]")) return name;
		if(jtype.equals("boolean[]")) return name;
		if(jtype.equals("String[]")) return name;
		return "@@@"+name;
	}
	/**
		return文を生成する
	*/
	String generateReturnStatement(ITypeInfo.ElemDesc elem, String var) {
		String jtype = getJavaType(elem.getTypeDesc());
		if(jtype.equals("void")) return "return;";			// 推奨されない
		if(jtype.equals("byte")) return "return ((Byte)"+var+").byteValue();";
		if(jtype.equals("short")) return "return ((Short)"+var+").shortValue();";
		if(jtype.equals("int")) return "return ((Integer)"+var+").intValue();";
		if(jtype.equals("float")) return "return ((Float)"+var+").floatValue();";
		if(jtype.equals("double")) return "return ((Double)"+var+").doubleValue();";
		if(jtype.equals("boolean")) return "return ((Boolean)"+var+").booleanValue();";
		if(jtype.equals("Object")) return "return "+var+";";			// 推奨されない
		if(jtype.equals("String")) return "return (String)"+var+";";
		if(jtype.equals("java.util.Date")) return "return (java.util.Date)"+var+";";
		if(elem.getTypeDesc().startsWith("VT_PTR+VT_USERDEFINED"))
				return "return new "+getJavaType(elem.getTypeDesc())+"((IDispatch)"+var+");";
		if(elem.getTypeDesc().equals("VT_DISPATCH"))
				return "return (IDispatch)"+var+";";
		if(elem.getTypeDesc().equals("VT_UNKNOWN"))
				return "return (IUnknown)"+var+";";

		return "return "+jtype+"@@@"+var;
	}

	//	ユーザ定義型の場合、その身元（ITypeLib）を洗う。
	public void checkUserDefined(ITypeInfo info, String type) {
		if(! type.startsWith("VT_PTR+VT_USERDEFINED")) return;
		try {
			// "VT_PTR+VT_USERDEFINED(50383648):Application"からhreftypeを取り出す
			int hreftype = ITypeInfo.getRefTypeFromTypeDesc(type);
			ITypeInfo refinfo = info.getRefTypeInfo(hreftype);
			ITypeLib lib = refinfo.getTypeLib();
			addTLibAttr(lib.getTLibAttr());
		}
		catch(Exception ex) { ex.printStackTrace(); }
	}
	// TLibAttrを集める
	Vector libset = new Vector();
	public boolean addTLibAttr(ITypeLib.TLibAttr attr) {
		// すでに同じITypeLibがあるかどうか？
		for(int i=0; i<libset.size(); i++) {
			if(libset.get(i).equals(attr)) return false;
		}
		libset.add(attr);
		return true;
	}

	// タイプライブラリから指定したITypeInfoの内容を出力する
	void generate(ITypeLib lib, String packagename, String infoname) {
		currentLib = lib;
		currentPackageName = packagename;
		try {
			int infocount = lib.getTypeInfoCount();
			// 指定したオブジェクトを表示。
			for(int i=0; i<infocount; i++) {
				ITypeInfo info = lib.getTypeInfo(i);
				if(info.getDocumentation(-1)[0].equals(infoname)) {
					generate(info);
					System.out.println("出力しました。多分");
					break;	// 終了
				}
			}
		} catch(Exception e) { e.printStackTrace(); }
	}

	// タイプライブラリの内容をすべて出力する
	void generateAll(ITypeLib lib, String packagename) {
		currentLib = lib;
		currentPackageName = packagename;
		try {
			// すべてのコンポーネントを表示。
			int infocount = lib.getTypeInfoCount();
			for(int i=0; i<infocount; i++) {
				ITypeInfo info = lib.getTypeInfo(i);
System.out.print(packagename+"."+info.getDocumentation(-1)[0]);
				generate(info);
System.out.println();
			}
System.out.println(""+infocount+" file(s)");
		} catch(Exception e) { e.printStackTrace(); }
	}

	/**
		使い方
	*/
	public static void usage() {
		System.out.println("指定したProgIDのオブジェクトのソースを自動生成します。");
		System.out.println();
		System.out.println("使い方: JComGenerator [<options>] <ProgID>");
		System.out.println("options");
		System.out.println("  -package <packagename>   パッケージにします");
		System.out.println("  -d <directory>           生成されるソースファイルの位置を指定します(末尾に\\なし)");
		System.out.println();
		System.out.println("例: JComGenerator \"Excel.Application\"");
		System.out.println("    すべてのオブジェクトのソースを生成します。");
		System.out.println("例: JComGenerator \"Excel.Application.8\" jp.ne.so_net.ga2.no_ji.jcom.excel8");
		System.out.println("    _Applicationオブジェクトのソースを生成します。");
		System.out.println();
		System.out.println("代表的な ProgID と Interface");
		System.out.println(" Excel.Application _Application");
		System.out.println(" InternetExplorer.Application IWebBrowser");
	}
	// とりあえず、ライブラリ名をくっつける。
	// jp.ne.so_net.ga2.no_ji.jcom.excelという具合になる
	String getPackageName(String packagepath, ITypeLib lib) throws JComException {
		if(packagepath==null) return lib.getDocumentation(-1)[0];
		return packagepath+"."+lib.getDocumentation(-1)[0];
	}

	public static void main(String[] args) throws Exception {
		// 引数解析
		if(args.length==0) { JComGenerator.usage(); return; }
		String packagepath = null;
		String outdirs = ".";
		String progID = null;
		for(int i=0; i<args.length; i++) {
			if(args[i].equals("-package")) {
				packagepath = args[++i];	// "jp.ne.so_net.ga2.no_ji.jcom"
			}
			else if(args[i].equals("-d")) {
				outdirs = args[++i];		// "../.."
			}
			else {
				progID = args[i];
			}
		}
		if(progID==null) { JComGenerator.usage(); return; }
		// 作る！
		JComGenerator generator = new JComGenerator();
		generator.setOutDirs(outdirs);	// 末尾に\なし
		generator.setPackagePath(packagepath);

		ReleaseManager rm = new ReleaseManager();
		try {
			System.out.println(progID+"を起動中...");
			IDispatch disp = new IDispatch(rm, progID);
			generator.progID = progID;
			// 指定されたProgIDを持つITypeLibを取得
			ITypeInfo typeInfo = disp.getTypeInfo();
			generator.entryInfo = typeInfo;
			ITypeLib typeLib = typeInfo.getTypeLib();

			// リストに追加。
			generator.addTLibAttr(typeLib.getTLibAttr());

			// ITypeLIbの内容をすべて解析して出力。
			// その際に、別のITypeLibのコンポーネントは随時溜めておいて、
			// それがなくなるまで出力する。
			for(int i=0; i<generator.libset.size(); i++) {
				ITypeLib.TLibAttr attr = (ITypeLib.TLibAttr)generator.libset.get(i);
System.out.println("TyepLib: "+attr);
				ITypeLib lib = ITypeLib.loadRegTypeLib(rm, attr.getLIBID(), attr.getVerMajor(), attr.getVerMinor());
				// パッケージ名を作る。
				String packagename = generator.getPackageName(packagepath, lib);
				// 生成するぜ！
				generator.generateAll(lib, packagename);
			}
		}
		catch(Exception e) { e.printStackTrace(); }
		finally { rm.release(); }
	}
}

/*
	generate(ITypeInfo info)
		generateEnum(ITypeInfo info)
		generateDispatch(ITypeInfo info)
		generateCoClass(ITypeInfo info)

*/

