package jp.ne.so_net.ga2.no_ji.jcom;

/**
	ITypeLibインターフェースを扱うためのクラス
	
	@see     ITypeInfo
	@see     IUnknown
	@see     JComException
	@see     ReleaseManager
	@author Yoshinori Watanabe(渡辺 義則)
	@version 2.10, 2000/07/01
	Copyright(C) Yoshinori Watanabe 1999-2000. All Rights Reserved.
 */
public class ITypeLib extends IUnknown {
    /**
		IID_ITypeLib です。
		@see       GUID
	*/
	public static GUID IID = new GUID( 0x00020402, 0x0000, 0x0000, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46 );

    /**
     * ITypeLibを作成します。
	 * 引数pITypeLibはITypeLibインターフェースのポインタを指定します。
     * @param     rm             参照カウンタ管理クラス
     	@param	pITypeLib	ITypeLibインターフェース
     * @see       ReleaseManager
	 */
	public ITypeLib(ReleaseManager rm, int pITypeLib) {
		super(rm, pITypeLib);
	}

	/**
		指定したメンバIDのドキュメントを返します。
		-1を指定した場合はこのオブジェクトに対するドキュメントを返します。
		戻り値[0]	bstrName, 
		戻り値[1]	btrDocString, 
		戻り値[2]	dwHelpContext, 
		戻り値[3]	bstrHelpFile, 
	*/
	public String[] getDocumentation(int index) throws JComException {
		return _getDocumentation(index);
	}
	public int getTypeInfoCount() throws JComException {
		return _getTypeInfoCount();
	}
	public ITypeInfo getTypeInfo(int index) throws JComException {
		return new ITypeInfo(rm, _getTypeInfo(index));
	}
	public TLibAttr getTLibAttr() throws JComException {
		return _getTLibAttr();
	}

	/**
		ITypeLIbの属性を管理するクラスです。
		LIBID、バージョン情報を提供します。
		変数は通常は定数で、プロパティとは異なります。
		プロパティの設定、取得は関数に含まれます。
		@see	ITypeLib
	*/
	public class TLibAttr {

		private GUID LIBID;
		private int verMajor;
		private int verMinor;
		public TLibAttr(GUID LIBID, int verMajor, int verMinor) {
			this.LIBID      = LIBID;
			this.verMajor   = verMajor;
			this.verMinor   = verMinor;
		}
		public GUID getLIBID() { return LIBID; }
		public int getVerMajor() { return verMajor; }
		public int getVerMinor() { return verMinor; }
		public String toString() { return LIBID.toString()+verMajor+"."+verMinor; }
		public boolean equals(Object obj) {
			if(! (obj instanceof TLibAttr)) return false;
			TLibAttr attr = (TLibAttr)obj;
			if(! attr.getLIBID().equals(LIBID)) return false;
			if(attr.getVerMajor() != verMajor) return false;
			if(attr.getVerMinor() != verMinor) return false;
			return true;
		}
	}

	// 静的メンバ
	/**
		タイプライブラリからITypeLibを取得します。
		どのファイルを読めばいいんだ？
		動作未確認
	*/
	public static ITypeLib loadTypeLib(ReleaseManager rm, String szFile) throws JComException {
		int pITypeLib = _loadTypeLibEx(szFile);
		if(pITypeLib == 0) return null;
		return new ITypeLib(rm, pITypeLib);
	}

	/**
		GUID形式のLIBIDとバージョン番号からITypeLibを取得します。
		
	*/
	public static ITypeLib loadRegTypeLib(ReleaseManager rm, GUID libid, int verMajor, int verMinor) throws JComException {
		int pITypeLib = _loadRegTypeLib(libid, verMajor, verMinor);
		if(pITypeLib == 0) return null;
		return new ITypeLib(rm, pITypeLib);
	}


	// release()はsuperのでＯＫ．

	// JNI
	private native String[]   _getDocumentation(int index) throws JComException;
	private native int        _getTypeInfo(int index) throws JComException;
	private native int        _getTypeInfoCount() throws JComException;
	private native TLibAttr   _getTLibAttr() throws JComException;
	private static native int _loadTypeLibEx(String szFile) throws JComException;
	private static native int _loadRegTypeLib(GUID guid, int verMajor, int verMinor) throws JComException;
	// jcom.dllを読み込みます。
    static {
        System.loadLibrary("jcom");
    }
}


