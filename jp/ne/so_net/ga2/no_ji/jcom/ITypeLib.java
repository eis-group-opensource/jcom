package jp.ne.so_net.ga2.no_ji.jcom;

/**
	ITypeLib�C���^�[�t�F�[�X���������߂̃N���X
	
	@see     ITypeInfo
	@see     IUnknown
	@see     JComException
	@see     ReleaseManager
	@author Yoshinori Watanabe(�n�� �`��)
	@version 2.10, 2000/07/01
	Copyright(C) Yoshinori Watanabe 1999-2000. All Rights Reserved.
 */
public class ITypeLib extends IUnknown {
    /**
		IID_ITypeLib �ł��B
		@see       GUID
	*/
	public static GUID IID = new GUID( 0x00020402, 0x0000, 0x0000, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46 );

    /**
     * ITypeLib���쐬���܂��B
	 * ����pITypeLib��ITypeLib�C���^�[�t�F�[�X�̃|�C���^���w�肵�܂��B
     * @param     rm             �Q�ƃJ�E���^�Ǘ��N���X
     	@param	pITypeLib	ITypeLib�C���^�[�t�F�[�X
     * @see       ReleaseManager
	 */
	public ITypeLib(ReleaseManager rm, int pITypeLib) {
		super(rm, pITypeLib);
	}

	/**
		�w�肵�������oID�̃h�L�������g��Ԃ��܂��B
		-1���w�肵���ꍇ�͂��̃I�u�W�F�N�g�ɑ΂���h�L�������g��Ԃ��܂��B
		�߂�l[0]	bstrName, 
		�߂�l[1]	btrDocString, 
		�߂�l[2]	dwHelpContext, 
		�߂�l[3]	bstrHelpFile, 
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
		ITypeLIb�̑������Ǘ�����N���X�ł��B
		LIBID�A�o�[�W��������񋟂��܂��B
		�ϐ��͒ʏ�͒萔�ŁA�v���p�e�B�Ƃ͈قȂ�܂��B
		�v���p�e�B�̐ݒ�A�擾�͊֐��Ɋ܂܂�܂��B
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

	// �ÓI�����o
	/**
		�^�C�v���C�u��������ITypeLib���擾���܂��B
		�ǂ̃t�@�C����ǂ߂΂����񂾁H
		���얢�m�F
	*/
	public static ITypeLib loadTypeLib(ReleaseManager rm, String szFile) throws JComException {
		int pITypeLib = _loadTypeLibEx(szFile);
		if(pITypeLib == 0) return null;
		return new ITypeLib(rm, pITypeLib);
	}

	/**
		GUID�`����LIBID�ƃo�[�W�����ԍ�����ITypeLib���擾���܂��B
		
	*/
	public static ITypeLib loadRegTypeLib(ReleaseManager rm, GUID libid, int verMajor, int verMinor) throws JComException {
		int pITypeLib = _loadRegTypeLib(libid, verMajor, verMinor);
		if(pITypeLib == 0) return null;
		return new ITypeLib(rm, pITypeLib);
	}


	// release()��super�̂łn�j�D

	// JNI
	private native String[]   _getDocumentation(int index) throws JComException;
	private native int        _getTypeInfo(int index) throws JComException;
	private native int        _getTypeInfoCount() throws JComException;
	private native TLibAttr   _getTLibAttr() throws JComException;
	private static native int _loadTypeLibEx(String szFile) throws JComException;
	private static native int _loadRegTypeLib(GUID guid, int verMajor, int verMinor) throws JComException;
	// jcom.dll��ǂݍ��݂܂��B
    static {
        System.loadLibrary("jcom");
    }
}


