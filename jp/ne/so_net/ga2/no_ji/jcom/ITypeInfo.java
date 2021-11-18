package jp.ne.so_net.ga2.no_ji.jcom;

/**
 * ITypeInfo�C���^�[�t�F�[�X���������߂̃N���X
 *
 * @see     IUnknown
 * @see     JComException
 * @see     ReleaseManager
	@author Yoshinori Watanabe(�n�� �`��)
	@version 2.10, 2000/07/01
	Copyright(C) Yoshinori Watanabe 1999-2000. All Rights Reserved.
 */
public class ITypeInfo extends IUnknown {
    /**
		IID_ITypeInfo �ł��B
		@see       GUID
	*/
	public static GUID IID = new GUID( 0x00020401, 0x0000, 0x0000, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46 );
    /**
     * ITypeInfo���쐬���܂��B
	 * ����pITypeInfo��ITypeInfo�C���^�[�t�F�[�X�̃|�C���^���w�肵�܂��B
     * @param     rm             �Q�ƃJ�E���^�Ǘ��N���X
     	@param	pITypeInfo	ITypeInfo�C���^�[�t�F�[�X
     * @see       ReleaseManager
	 */
	public ITypeInfo(ReleaseManager rm, int pITypeInfo) {
		super(rm, pITypeInfo);
	}
	/**
		�w�肵�������oID�̃h�L�������g��Ԃ��܂��B
		�����o�h�c��FuncDesc�N���X��memid�ɂ���܂��B
		MEMBERID_NIL(-1)���w�肵���ꍇ�͂��̃I�u�W�F�N�g�ɑ΂���h�L�������g��Ԃ��܂��B
		�߂�l�͂S���̕�����̔z��ŁA���ꂼ��ȉ��̒l���i�[����Ă��܂��B
		�J�b�R����Excel.Application�̏ꍇ�̒l�ł��B
		�l�̂Ȃ����̂�null�ɂȂ��Ă��܂��B
		<pre>
		�߂�l[0]	bstrName        ���̖��O�B("_Application")
		�߂�l[1]	btrDocString    �h�L�������g(null)
		�߂�l[2]	dwHelpContext    �w���v�R���e�L�X�g�̔ԍ��𕶎���ɕς�������("131073")
		�߂�l[3]	bstrHelpFile	�w���v�t�@�C���̃t���p�X("D:\Office97\Office\VBAXL8.HLP")
		<pre>
	*/
	public String[] getDocumentation(int memid) throws JComException {
		return _getDocumentation(memid);
	}
	public static final int MEMBERID_NIL = -1;
	/**
		TypeAttr��Ԃ��܂��B
		TypeAttr��ITypeInfo�̑������Ǘ�����N���X�ł��B
		@see	ITypeInfo.TypeAttr
	*/
	public TypeAttr getTypeAttr() throws JComException {
		return _getTypeAttr();
	}
	/**
		�w�肵��index��FuncDesc��Ԃ��܂��B
		FuncDesc�̓��\�b�h�̏����Ǘ�����N���X�ł��B
		@see	ITypeInfo.FuncDesc
	*/
	public FuncDesc getFuncDesc(int index) throws JComException {
		return _getFuncDesc(index);
	}
	/**
		�w�肵��index��VarDesc��Ԃ��܂��B
		VarDesc�͕ϐ�(�ʏ��Enum�^)�̏����Ǘ�����N���X�ł��B
		@see	ITypeInfo.VarDesc
	*/
	public VarDesc getVarDesc(int index) throws JComException {
		return _getVarDesc(index);
	}
	/**
		���̃I�u�W�F�N�g�������Ă���ITypeLib��Ԃ��܂��B
		@see	ITypeLib
	*/
	public ITypeLib getTypeLib() throws JComException {
		return new ITypeLib(rm, _getTypeLib());
	}
	/**
		���̃N���X��COCLASS�̂Ƃ��A���ۂɎ������Ă���^����Ԃ��܂��B
	*/
	public ITypeInfo getImplType(int index) throws JComException {
		return new ITypeInfo(rm, _getImplType(index));
	}
	/**
		hreftype�Ŏw�肳�ꂽITypeInfo��Ԃ��܂��B
		@see ITypeInfo.ElemDesc#getTypeDesc()
	*/
	public ITypeInfo getRefTypeInfo(int hreftype) throws JComException {
		return new ITypeInfo(rm, _getRefTypeInfo(hreftype));
	}
	/**
		�^���VT_USERDEFINED�̂Ƃ��A��������HREFTYPE�����o���܂��B
		HREFTYPE��16�i���ŕ\����Ă��܂��B
	*/
	public static int getRefTypeFromTypeDesc(String type) {
		return Integer.parseInt(type.substring(type.indexOf('(')+1, type.indexOf(')')),16);
	}
	/**
		���̃I�u�W�F�N�g�Ƒ��̃I�u�W�F�N�g�����������ǂ����������܂��B 
		ITypeInfo.TypeAttr.IID �ɂ�蓯��̌^��񂩂ǂ����𔻒f���Ă��܂��B
		@see ITypeInfo.TypeAttr#getIID()
	*/
	public boolean equals(Object obj) {
		if(! (obj instanceof ITypeInfo)) return false;
		ITypeInfo info = (ITypeInfo)obj;
		try {
			return info.getTypeAttr().getIID().equals(this.getTypeAttr().getIID());
		} catch(JComException e) { e.printStackTrace(); }
		return false;
	}
	/**
		ITypeInfo�̑������Ǘ�����N���X�ł��B
		GUID�A�֐��̐��A�ϐ��̐���Ԃ��܂��B
		�ϐ��͒ʏ�͒萔�ŁA�v���p�e�B�Ƃ͈قȂ�܂��B
		�v���p�e�B�̐ݒ�A�擾�͊֐��Ɋ܂܂�܂��B
		@see	ITypeInfo
	*/
	public class TypeAttr {
		public static final int TKIND_ENUM      = 0;
		public static final int TKIND_RECORD    = 1;
		public static final int TKIND_MODULE    = 2;
		public static final int TKIND_INTERFACE = 3;
		public static final int TKIND_DISPATCH  = 4;
		public static final int TKIND_COCLASS   = 5;
		public static final int TKIND_ALIAS     = 6;
		public static final int TKIND_UNION     = 7;
		public static final int TKIND_MAX       = 8;

		private GUID IID;
		private int cFuncs;
		private int cVars;
		private int cImplTypes;
		private int typekind;
		public TypeAttr(GUID IID, int typekind, int cFuncs, int cVars, int cImplTypes) {
			this.IID        = IID;
			this.typekind   = typekind;
			this.cFuncs     = cFuncs;
			this.cVars      = cVars;
			this.cImplTypes = cImplTypes;
		}
		public GUID getIID() { return IID; }
		public int getTypeKind() { return typekind; }
		public int getFuncs() { return cFuncs; }
		public int getVars() { return cVars; }
		public int getImplTypes() { return cImplTypes; }
	}


	/**
		�P�̈�����߂�l�Ȃǂ̌^�̏��������܂��B
		�b�n�l��ELEMDESC�\���̂�\���N���X�ł��B
		[in out]�Ȃǂ̏��ƁA�^ VT_PTR+VT_BSTR �Ȃǂ̏��������܂��B
		VT_USERDEFINED�����ꍇ�A�����I�ɂ��̃N���X�����擾���A
		���ۂ̕�����ɒu�������܂��B
		�^�͕�����̌`�ŕێ����AJava�݊��̌`�ɂȂ�܂��B���Ȃ킿�A
		VT_PTR|VT_I4�� "[I"�AVT_UNKNOWN��"Ljp.ne.so_net.ga2.no_ji.jcom.IUnknown"��
		�Ȃ�܂��B
		@see	ITypeInfo.FuncDesc
	*/
	public class ElemDesc {
		public static final int IDLFLAG_FIN     = 0x01;
		public static final int IDLFLAG_FOUT    = 0x02;
		public static final int IDLFLAG_FLCID   = 0x04;
		public static final int IDLFLAG_FRETVAL = 0x08;
		private int idl;	// IDLFLAG_XXX�̑g�ݍ��킹
		private String typedesc;	// �^��� "VT_INT"
		public ElemDesc(String typedesc, int idl) {
			this.typedesc = typedesc;
			this.idl = idl;
		}
		public int getIDL() { return idl; }
		/**
			�^����Ԃ��܂��B�ȉ��̌`���ł��B
			"VT_I4" "VT_BSTR" "VT_DISPATCH" "VT_PTR+VT_I4"
			"VT_SAFEARRAY+VT_I4"
			"VT_USERDEFINED(1):VBE"
			VT_USERDEFINED�̊��ʓ��̐��l��16�i���ŕ\������hreftype�ŁA���̐��l��
			getRefTypeInfo()�ɓn�����Ƃɂ��A���̒l��ITypeInfo��
			�擾���邱�Ƃ��ł��܂��B
		*/
		public String getTypeDesc() { return typedesc; }
		public String toString() {
			if(idl==0) return typedesc;
			String result = "";
			if((idl & IDLFLAG_FIN)!=0) result += "[in]";
			if((idl & IDLFLAG_FOUT)!=0) result += "[out]";
			if((idl & IDLFLAG_FLCID)!=0) result += "[lcid]";	// ???
			if((idl & IDLFLAG_FRETVAL)!=0) result += "[retval]";
			return result + typedesc;
		}
	}
	
	/**
		���\�b�h�̏����Ǘ�����N���X�ł��B
		�����o�h�c�A����ьĂяo���`���A�����̌^��߂�l�̌^�Ȃǂ�
		�����܂��B
		@see	ITypeInfo.ElemDesc
		@see	ITypeInfo
	*/
	public class FuncDesc {
		private int memid;					// �����oID(0�`)
		private int invkind;				// INVOKE_XXX�̂����ꂩ
		private ElemDesc[] elemdescParam;	// �����̌^
		private ElemDesc elemdescFunc;		// �߂�l�̌^
		public static final int INVOKE_FUNC           = 0x01;
		public static final int INVOKE_PROPERTYGET    = 0x02;
		public static final int INVOKE_PROPERTYPUT    = 0x04;
		public static final int INVOKE_PROPERTYPUTREF = 0x08;
		/**
			���\�b�h�̏��𐶐����܂��B
			ITypeInfo.getFuncDesc()���Ŏg�p����܂��B
			�ʏ�A�O������͎g�p���܂���B
			@see	ITypeInfo#getFuncDesc(int)
		*/
		public FuncDesc(int memid, int invkind, ElemDesc[] elemdescParam, ElemDesc elemdescFunc) {
			this.memid = memid;
			this.invkind = invkind;
			this.elemdescParam = elemdescParam;
			this.elemdescFunc = elemdescFunc;
		}
		/**
			���\�b�h�̏���\�����܂��B
		*/
		public String toString() {
			try {
				String[] names = getNames();
				String result = "";
				switch(invkind) {
					case INVOKE_FUNC:			result += "FUNC ";
						break;
					case INVOKE_PROPERTYGET:	result += "GET ";	break;
					case INVOKE_PROPERTYPUT:	result += "PUT ";	break;
					case INVOKE_PROPERTYPUTREF:	result += "PUTREF ";	break;
				}
				result += names[0] + "(";
				if(elemdescParam!=null) {
					for(int i=0; i<elemdescParam.length; i++) {
						result += elemdescParam[i].toString() + " ";
						if(i+1<names.length) result += names[i+1];
						if(i!=elemdescParam.length-1) result += ",";
					}
				}
				result += ")"+elemdescFunc.toString();
				return result;
			}
			catch(Exception e) {}
			return null;
/*
			try {
				String[] names = getNames();
				String result = "";
				switch(invkind) {
					case INVOKE_FUNC:			result += "FUNC ";	break;
					case INVOKE_PROPERTYGET:	result += "GET ";	break;
					case INVOKE_PROPERTYPUT:	result += "PUT ";	break;
					case INVOKE_PROPERTYPUTREF:	result += "PUTREF ";	break;
				}
				result += names[0] + "(";
				if(elemdescParam!=null) {
					for(int i=0; i<elemdescParam.length; i++) {
						result += elemdescParam[i].toString() + " " + names[i+1];
						if(i!=elemdescParam.length-1) result += ",";
					}
				}
				result += ")"+elemdescFunc.toString();
				return result;
			}
			catch(Exception e) {}
			return null;
*/
		}
		/**
			�����o�h�c��Ԃ��܂��B0�ȏ�̒l�ł��B
		*/
		public int getMemID() { return memid; }
		/**
			�Ăяo���`����Ԃ��܂��B
			INVOKE_XXX�̂����ꂩ�ł��B
			@see	ITypeInfo.FuncDesc#INVOKE_FUNC
			@see	ITypeInfo.FuncDesc#INVOKE_PROPERTYGET
			@see	ITypeInfo.FuncDesc#INVOKE_PROPERTYPUT
			@see	ITypeInfo.FuncDesc#INVOKE_PROPERTYPUTREF
		*/
		public int getInvokeKind() { return invkind; }
		/**
			�����̏���Ԃ��܂��B
			�������Ȃ��ꍇ��<code>null</code>��Ԃ��܂��B
			@see	ITypeInfo.ElemDesc
		*/
		public ElemDesc[] getParams() { return elemdescParam; }
		/**
			�߂�l�̏���Ԃ��܂��B
			@see	ITypeInfo.ElemDesc
		*/
		public ElemDesc getFunc() { return elemdescFunc; }
		/**
			���\�b�h�̖��O�A�����̖��O��Ԃ��܂��B
			�ǂ����A[0]�����\�b�h���A[1]�ȍ~�������̖��O�̂悤�ł��B
		*/
		public String[] getNames() throws JComException {
			return _getNames(memid);
		}
	}

	/**
		�����ϐ��̏����Ǘ�����N���X�ł��B
		�ʏ�AEnum�^�̒萔�Ɏg���܂��B
		@see	ITypeInfo
		@see	ITypeInfo.ElemDesc
	*/
	public class VarDesc {
		private int memid;					// �����oID(0�`)
		private int varkind;				// VAR_XXX�̂����ꂩ
		private ElemDesc elemdescVar;		// �ϐ��̌^
		private Object   varValue;			// �߂�l�̌^
		public static final int VAR_PERINSTANCE = 0;
		public static final int VAR_STATIC      = 1;
		public static final int VAR_CONST       = 2;
		public static final int VAR_DISPATCH    = 3;
		/**
			���\�b�h�̏��𐶐����܂��B
			ITypeInfo.getFuncDesc()���Ŏg�p����܂��B
			�ʏ�A�O������͎g�p���܂���B
			@see	ITypeInfo#getFuncDesc(int)
		*/
		public VarDesc(int memid, int varkind, ElemDesc elemdescVar, Object varValue) {
			this.memid = memid;
			this.varkind = varkind;
			this.elemdescVar = elemdescVar;
			this.varValue = varValue;
		}
		/**
			���\�b�h�̏���\�����܂��B
		*/
		public String toString() {
			try {
				String[] names = getNames();
				String result = "";
				switch(varkind) {
					case VAR_PERINSTANCE:	result += "PERINSTANCE ";	break;
					case VAR_STATIC:		result += "STATIC ";		break;
					case VAR_CONST:			result += "CONST ";			break;
					case VAR_DISPATCH:		result += "DISPATCH ";		break;
				}
				result += elemdescVar.toString() + " " + names[0] + " = " + varValue.toString();
				return result;
			}
			catch(Exception e) {}
			return null;
		}
		/**
			�����o�h�c��Ԃ��܂��B0�ȏ�̒l�ł��B
		*/
		public int getMemID() { return memid; }
		/**
			�ϐ��̌`����Ԃ��܂��B
			VAR_XXX�̂����ꂩ�ł��B
			@see	ITypeInfo.VarDesc#VAR_PERINSTANCE
			@see	ITypeInfo.VarDesc#VAR_STATIC
			@see	ITypeInfo.VarDesc#VAR_CONST
			@see	ITypeInfo.VarDesc#VAR_DISPATCH
		*/
		public int getVarKind() { return varkind; }
		/**
			�ϐ��̌^�̏���Ԃ��܂��B
			@see	ITypeInfo.ElemDesc
		*/
		public ElemDesc getVar() { return elemdescVar; }
		/**
			���\�b�h�̖��O�A�����̖��O��Ԃ��܂��B
			�ǂ����A[0]�����\�b�h���A[1]�ȍ~�������̖��O�̂悤�ł��B
		*/
		public String[] getNames() throws JComException {
			return _getNames(memid);
		}
		/**
			�ϐ��̒l��Ԃ��܂��B
		*/
		public Object getValue() { return varValue; }
	}

	// release()��super�̂łn�j�D

	// JNI
	private native String[] _getDocumentation(int memid) throws JComException;
	private native String[] _getNames(int memid) throws JComException;
	private native TypeAttr _getTypeAttr() throws JComException;
	private native FuncDesc _getFuncDesc(int index) throws JComException;
	private native VarDesc  _getVarDesc(int index) throws JComException;
	private native int      _getImplType(int index) throws JComException;
	private native int      _getTypeLib() throws JComException;		// ITypeInfo::GetContainingTypeLib()
	private native int      _getRefTypeInfo(int hreftype) throws JComException;		// ITypeInfo::GetRefTypeInfo()

}

/*
	FuncDesc �� ElemDesc �́A�ꉞ�A�O������ύX���ł��Ȃ��悤�ɂ��Ă���B
	���Ȃ킿�A�����o��private�ɂ��Aget�n�����Ή����Ă��Ȃ��B
	�����A�ǂ�����R���X�g���N�^��public�Ȃ̂Ŋ��S�Ƃ͌����Ȃ��B
	�i�m�h�����Ő�������K�v�����邩��ł��B
	�f�U�C���p�^�[����singleton�̍l���������Ă��܂��B

	�������A���N�̉Ă������`�B
2000-07-23
	COCLASS���ĂȂ񂶂Ⴀ�`�I�I�I�I

*/
