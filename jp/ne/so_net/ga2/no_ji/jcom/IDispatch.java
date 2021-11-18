package jp.ne.so_net.ga2.no_ji.jcom;

/**
 * IDispatch�C���^�[�t�F�[�X���������߂̃N���X
 *
 * @see     IUnknown
 * @see     JComException
 * @see     ReleaseManager
	@author Yoshinori Watanabe(�n�� �`��)
	@version 2.10, 2000/08/23
	Copyright(C) Yoshinori Watanabe 1999-2000. All Rights Reserved.
 */
public class IDispatch extends IUnknown {
    /**
		IID_IDispatch �ł��B
		@see       GUID
	*/
	public static GUID IID = new GUID( 0x00020400, 0x0000, 0x0000, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46 );
	public static final int METHOD         = 1;
	public static final int PROPERTYGET    = 2;
	public static final int PROPERTYPUT    = 4;
	public static final int PROPERTYPUTREF = 8;

    /**
     * IDispatch���쐬���܂��B
     * �����Ŏw�肳�ꂽProgID��COM�I�u�W�F�N�g���쐬���܂��B
     *        <pre>
     *        ReleaseManager rm = new ReleaseManager();
     *        try {
     *            IDispatch excel = new IDispatch(rm ,"Excel.Application");
     *            excel.put("Visible", new Boolean(true));  // '�f�t�H���g��False(�\�����Ȃ�)
     *            // ...
     *            excel.invoke("Quit", null);
     *        } catch(JComException e) {
     *            e.printStackTrace();
     *        } finally {
     *            rm.release();
     *        }</pre>
     * @param     rm     	�Q�ƃJ�E���^�Ǘ��N���X
     * @param	  ProgID	�v���O�����h�c�BExcel�̏ꍇ"Excel.Application"�Ǝw�肵�܂��B
     * @see       #create(String)
     * @see       ReleaseManager
	 */
	public IDispatch(ReleaseManager rm, String ProgID) throws JComException {
		super(rm);
		create(ProgID);
	}

    /**
     * IDispatch���쐬���܂��B
     * �����Ŏw�肳�ꂽIDispatch�|�C���^��COM���쐬���܂��B
     * �ʏ�A�ʂ�COM�̃��\�b�h����Ԃ��ꂽIDispatch�|�C���^�ɑ΂��āA
     * �g�p���܂��B
     * @param     rm             �Q�ƃJ�E���^�Ǘ��N���X
     * @param     pIDispatch     IDispatch�C���^�[�t�F�[�X�̃A�h���X
     * @see       #create(String)
     * @see       ReleaseManager
	 */
	public IDispatch(ReleaseManager rm, int pIDispatch) {
		super(rm, pIDispatch);
	}

    /**
     * IDispatch���쐬���܂��B
     * �����Ŏw�肳�ꂽIDispatch�Ɠ���̂b�n�l���Ǘ����܂��B
     * �ʏ�A�ʂ�COM�I�u�W�F�N�g�̃��\�b�h�ŕԂ��ꂽIDispatch�ɑ΂��āA
     * �g�p���܂��B
     * ReleaseManager�́A����disp�̎��������̂��p������܂��B
     * @see       #create(String)
	 */
	public IDispatch(IDispatch disp) {
		super(disp.rm, disp.pIUnknown);
	}

    /**
     * ProgID����IDispatch�C���^�[�t�F�[�X���쐬���܂��B
     * ���łɍ쐬����Ă����ꍇ�A��O�𔭐������܂��B
     * @param     ProgID COM�̃v���O����ID
     * @exception JComException <BR>
     *              "Already COM allocated" ���ł�COM�����蓖�Ă��Ă���B<BR>
     *              "createInstance() failed HRESULT=0x%XL" �b�n�l�̍쐬�Ɏ��s�����B
     */
	public synchronized void create(String ProgID) throws JComException {
		_create(ProgID);
		if(rm!=null) rm.add(this);
	}

    /**
     * �v���p�e�B�̒l���擾���܂��B
     * �v���p�e�B�̌^�Ƃi�������̌^�Ƃ̑Ή��͈ȉ��̒ʂ�ł��B
     * <pre>
     *   VT_EMPTY    null
     *   VT_I4       Integer
     *   VT_R8       Double
     *   VT_BOOL     Boolean
     *   VT_BSTR     String
     *   VT_DATE     Date
     *   VT_CY       VariantCurrency
     *   VT_DISPATCH IDispatch
     *   VT_UNKNOWN  IUnknown
     * </pre>
     *	<pre>IDispatch xlBooks = (IDispatch)xlApp.get("Workbooks");</pre>
	 *  <pre>Boolean visible = xlApp.gut("Visible");</pre>
     * @param     property �v���p�e�B��
     * @return    �擾���ꂽ�l
     * @exception JComException <BR>
     *              "IDispatch not created"   IDispatch���쐬����Ă��Ȃ��B<BR>
     *              "getProperty() failed HRESULT=0x%XL" �b�n�l�̌Ăяo���Ɏ��s�����B<BR>
     *              "cannot convert VT=0x%X" ���Ή���VARIANT�^���Ԃ��ꂽ�B
     * @see       #get(String,Object[])
     * @see       #invoke(String,Object[])
     * @see       #put(String,Object)
     */
/*	public synchronized Object get(String property) throws JComException {
		Object ret = _get(property);
		if((rm!=null) && (ret instanceof IUnknown))
			rm.add((IUnknown)ret);
		return ret;
	}
*/
	public synchronized Object get(String property) throws JComException {
		int dispid = _getIDsOfNames(property);
		Object ret = _invoke(dispid, PROPERTYGET, null);
		if((rm!=null) && (ret instanceof IUnknown))
			rm.add((IUnknown)ret);
		return ret;
	}


    /**
     * �v���p�e�B�̒l���擾���܂��B�v���p�e�B�̒l�̎擾�Ɉ������K�v�ȏꍇ�Ɏg�p���܂��B
     * �v���p�e�B�̌^�Ƃi�������̌^�Ƃ̑Ή���IDispatch.get(property)���Q�Ƃ��Ă��������B
     * �����̓n������IDispatch.invoke()���Q�Ƃ��Ă��������B
     * @param     property �v���p�e�B��
     * @param     args �����̔z��
     * @return    �擾���ꂽ�l
     * @exception JComException <BR>
     *              "IDispatch not created"   IDispatch���쐬����Ă��Ȃ��B<BR>
     *              "Invalid argument(index=%d)" �������s���B�������͖��Ή��̌^���n���ꂽ�B<BR>
     *              "getPropertyArg() failed HRESULT=0x%XL" �b�n�l�̌Ăяo���Ɏ��s�����B<BR>
     *              "cannot convert VT=0x%X" ���Ή���VARIANT�^���Ԃ��ꂽ�B
     * @see       #get(String)
     * @see       #invoke(String,Object[])
     * @see       #put(String,Object)
     */
/*	public synchronized Object get(String property, Object[] args) throws JComException {
		Object ret = _get(property, args);
		// �b�n�l�I�u�W�F�N�g���Ԃ��ꂽ��A������Q�ƃJ�E���^�Ǘ��N���X�ɒǉ�����B
		if((rm!=null) && (ret instanceof IUnknown))
			rm.add((IUnknown)ret);
		return ret;
	}
*/
	public synchronized Object get(String property, Object[] args) throws JComException {
		int dispid = _getIDsOfNames(property);
		Object ret = _invoke(dispid, PROPERTYGET, args);
		// �b�n�l�I�u�W�F�N�g���Ԃ��ꂽ��A������Q�ƃJ�E���^�Ǘ��N���X�ɒǉ�����B
		if((rm!=null) && (ret instanceof IUnknown))
			rm.add((IUnknown)ret);
		return ret;
	}


    /**
     * �v���p�e�B�ɒl��ݒ肵�܂��B
     * �v���p�e�B�̌^�Ƃi�������̌^�Ƃ̑Ή���IDispatch.get(property)���Q�Ƃ��Ă��������B
			<pre>xlApp.put("Visible", new Boolean(true));</pre>
			<pre>xlRange.put("Value","JCom���������I(^o^)");</pre>
     * @param     property �v���p�e�B��
     * @param     val �ݒ肷��l
     * @exception JComException <BR>
     *              "IDispatch not created"   IDispatch���쐬����Ă��Ȃ��B<BR>
     *              "Invalid argument" �������s���B�������͖��Ή��̌^���n���ꂽ�B<BR>
     *              "putProperty() failed HRESULT=0x%XL" �b�n�l�̌Ăяo���Ɏ��s�����B<BR>
     * @see       #get(String)
     * @see       #get(String,Object[])
     * @see       #invoke(String,Object[])
     */
/*	public synchronized void   put(String property, Object val) throws JComException {
		_put(property, val);
	}
*/
	public synchronized void   put(String property, Object val) throws JComException {
		int dispid = _getIDsOfNames(property);
		Object[] args = new Object[1];
		args[0] = val;
		_invoke(dispid, PROPERTYPUT, args);
	}



    /**
     * ���\�b�h���Ăяo���܂��B
     * �v���p�e�B�̌^�Ƃi�������̌^�Ƃ̑Ή���JCom.get(property)���Q�Ƃ��Ă��������B
     * <pre>
			Object[] arglist = new Object[3];
			arglist[0] = new Boolean(false);
			arglist[1] = null;
			arglist[2] = new Boolean(false);
			xlBook.method("Close", arglist);
     * </pre>
     * @param     method ���\�b�h��
     * @param     args   ����
     * @exception JComException <BR>
     *              "IDispatch not created"   IDispatch���쐬����Ă��Ȃ��B<BR>
     *              "Invalid argument(index=%d)" �������s���B�������͖��Ή��̌^���n���ꂽ�B<BR>
     *              "invokeMethod() failed HRESULT=0x%XL" �b�n�l�̌Ăяo���Ɏ��s�����B<BR>
     *              "cannot convert VT=0x%X" ���Ή���VARIANT�^���Ԃ��ꂽ�B
     * @see       #get(String)
     * @see       #get(String,Object[])
     * @see       #put(String,Object)
	 */
/*
	public synchronized Object method(String method, Object[] args) throws JComException {
		Object ret = _method(method, args);
		if((rm!=null) && (ret instanceof IUnknown))
			rm.add((IUnknown)ret);
		return ret;
	}
*/
	public synchronized Object method(String method, Object[] args) throws JComException {
		int dispid = _getIDsOfNames(method);
		Object ret = _invoke(dispid, METHOD, args);
		if((rm!=null) && (ret instanceof IUnknown))
			rm.add((IUnknown)ret);
		return ret;
	}

    /**
		method()���Q�Ƃ��Ă��������B
		@see	#method(String,Object[])
		@deprecated	method(String,Object[])�ɒu�������܂����B
	 */
	public synchronized Object invoke(String method, Object[] args) throws JComException {
		return method(method, args);
	}

    /**
		���\�b�h�A�v���p�e�B�̐ݒ�E�擾���s���܂��B
		�v���p�e�B�̌^�Ƃi�������̌^�Ƃ̑Ή���JCom.get(property)���Q�Ƃ��Ă��������B
		<pre>
			Object[] arglist = new Object[3];
			arglist[0] = new Boolean(false);
			arglist[1] = null;
			arglist[2] = new Boolean(false);
			xlBook.invoke("Close", IDispatch.METHOD, arglist);
		</pre>
		@param     method ���\�b�h��
		@param     args   ����
		@exception JComException <BR>
		             "IDispatch not created"   IDispatch���쐬����Ă��Ȃ��B<BR>
		             "Invalid argument(index=%d)" �������s���B�������͖��Ή��̌^���n���ꂽ�B<BR>
		             "invokeMethod() failed HRESULT=0x%XL" �b�n�l�̌Ăяo���Ɏ��s�����B<BR>
		             "cannot convert VT=0x%X" ���Ή���VARIANT�^���Ԃ��ꂽ�B
		@see       #get(String)
		@see       #get(String,Object[])
		@see       #put(String,Object)
		@see       #method(String,Object[])
	*/
	public synchronized Object invoke(String name, int wFlags, Object[] pDispParams) throws JComException {
		int dispid = _getIDsOfNames(name);
		return _invoke(dispid, wFlags, pDispParams);
	}

    /**
    	ITypeInfo���擾���܂��B
	*/
	public synchronized ITypeInfo getTypeInfo() throws JComException {
		int pITypeInfo = _getTypeInfo();
		return new ITypeInfo(rm, pITypeInfo);
	}

	// release()��super�̂łn�j�D

	// JNI
	private native void   _create(String ProgID) throws JComException;
	private native Object _get(String property) throws JComException;
	private native Object _get(String property, Object[] args) throws JComException;
	private native void   _put(String property, Object val) throws JComException;
	private native Object _method(String method, Object[] args) throws JComException;
	private native int    _getTypeInfo() throws JComException;
	private native Object _invoke(int dispIdMember, int wFlags, Object[] pDispParams) throws JComException;
	private native int    _getIDsOfNames(String name) throws JComException;

	// jcom.dll��ǂݍ��݂܂��B
    static {
        System.loadLibrary("jcom");
    }
}

