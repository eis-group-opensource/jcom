package jp.ne.so_net.ga2.no_ji.jcom;
import java.lang.reflect.*;

/**
 * IUnknown�C���^�[�t�F�[�X���������߂̃N���X
 *
 * @see     IDispatch
 * @see     JComException
 * @see     ReleaseManager
	@author Yoshinori Watanabe(�n�� �`��)
	@version 2.00, 2000/06/25
	Copyright(C) Yoshinori Watanabe 1999-2000. All Rights Reserved.
 */
public class IUnknown {
    /**
		IID_IUnknown �ł��B
		@see       GUID
	*/
	public static GUID IID = new GUID( 0x00000000, 0x0000, 0x0000, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46 );

	/**
		IUnknown�C���^�[�t�F�[�X�̃|�C���^��ێ����܂��B
		�ύX���Ȃ��ł��������B
		IUnknown interface pointer.
		Don't change!
	*/
	protected int pIUnknown = 0;		// IUnknown�C���^�[�t�F�[�X�����I�u�W�F�N�g�̃|�C���^
	/**
		�Q�ƃJ�E���^�Ǘ��N���X
		Reference counter management class.
	*/
	protected ReleaseManager rm = null;

    /**
     * ���IUnknown���쐬���܂��B
	 * COM�͊��蓖�Ă��Ă��Ȃ��̂Œ���
     * @param     rm             �Q�ƃJ�E���^�Ǘ��N���X
     * @see       ReleaseManager
	 */
	public IUnknown(ReleaseManager rm) {
		this.rm = rm;
		this.pIUnknown = 0;
	}

    /**
     * IUnknown���쐬���܂��B
     * �����Ŏw�肳�ꂽIUnknown�|�C���^��COM���쐬���܂��B
     * �ʏ�A�ʂ�COM�̃��\�b�h����Ԃ��ꂽIUnknown�|�C���^�ɑ΂��āA
     * �g�p���܂��B
     * @param     rm             �Q�ƃJ�E���^�Ǘ��N���X
     * @param     IUnknown       IUnknown�C���^�[�t�F�[�X�̃A�h���X
     * @see       ReleaseManager
	 */
	public IUnknown(ReleaseManager rm, int pIUnknown) {
		this.rm = rm;
		this.pIUnknown = pIUnknown;
	}

    /**
     * QueryInterface�����s���A�w�肵���C���^�[�t�F�[�X���擾���܂��B
     * IID �� Java���̃N���X���w�肵�ĉ������B
     *  <pre>
     *  IUnknown iUnknown = (IUnknown)worksheets.get("_NewEnum");
     *  IEnumVARIANT a = (IEnumVARIANT)iUnknown.queryInterface(
     *                      IEnumVARIANT.class,
     *                      IEnumVARIANT.IID);
     *  </pre>
		2000.11.27 �T�|�[�g���Ă��Ȃ��C���^�[�t�F�[�X�͊m����null��Ԃ��悤�ɂ����B
     * @param     jclass         �i���������̃N���X
     * @param     IID            �C���^�[�t�F�[�X�h�c
     * @result    �w�肵���C���^�[�t�F�[�X��Ԃ��܂��B
     *            ���s�����null��Ԃ��܂��B
     * @see       GUID
	 */
	public synchronized IUnknown queryInterface(Class jclass, GUID IID) throws JComException {
		try {
			int pIUnknown = _queryInterface(IID);
			if(pIUnknown == 0) return null;		// no interface
			// �N���X���̃R���X�g���N�^���ĂԁB������(ReleaseManager rm, int pIUnknown);
			Class[] param = new Class[2];
			param[0] = ReleaseManager.class;
			param[1] = Integer.TYPE;
			Constructor constructor = jclass.getConstructor(param);
			Object[] p = new Object[2];
			p[0] = rm;
			p[1] = new Integer(pIUnknown);
			IUnknown com = (IUnknown)constructor.newInstance(p);
			// ReleaseManager�ɓo�^
			rm.add(com);
			return com;
		}
		// JComException�͏�ɓ�����B
		catch(JComException e) { throw e; }
		catch(Exception e) { e.printStackTrace(); }
		return null;	
	}

    /**
     * QueryInterface�����s���A�w�肵���C���^�[�t�F�[�X���擾���܂��B
     * IID �� Java���̃N���X�����w�肵�ĉ������B
     * @param     classname      �i���������̃N���X��
     * @param     IID            �C���^�[�t�F�[�X�h�c
     * @result    �w�肵���C���^�[�t�F�[�X��Ԃ��܂��B
     *            ���s�����null��Ԃ��܂��B
     * @see       GUID
     * @deprecation #queryInterface(Class, GUID)
	 */
	public synchronized IUnknown queryInterface(String classname, GUID IID)
					 throws JComException, ClassNotFoundException {
		Class jclass = Class.forName(classname);
		return queryInterface(jclass, IID);
	}

    /**
     * COM�I�u�W�F�N�g��������܂��B
     * ReleaseManager���g���΁A�����I�ɌĂԕK�v�͂���܂���B
     * ���łɉ������Ă����ꍇ�Afalse��Ԃ��܂��B
     * @result    ����I���̏ꍇ��<code> true </code>��Ԃ��܂��B
     *            ���łɉ������Ă����ꍇ�́A<code> false </code>��Ԃ��܂��B
	 */
	public synchronized boolean release() {
		return _release();
	}

    /**
     * �Q�ƃJ�E���^���P�����A���݂̃J�E���^�l��Ԃ��܂��B
     * �ʏ�ĂԕK�v�͂���܂���B
     * �Q�ƃJ�E���^���������ꍇ�ł̂ݎg�p���܂��B
     * ���̏ꍇ��release()���Ă�ŁA�J�E���^�������ĉ������B
     * @result    �Q�ƃJ�E���^�̒l�B�オ�����l��Ԃ��܂��B
     * @see       #release()
	 */
	public synchronized int addRef() {
		return _addRef();
	}
	
	/**
		�����ŕێ����Ă���IUnknown�N���X�A�܂��͂��ꂩ��p�������N���X��
		�I�u�W�F�N�g���ȉ��̌`�ŕ\�����܂��B
		<pre>476eb8(1)jp.ne.so_net.ga2.no_ji.jcom.IDispatch</pre>
		�P�U�i���̓C���^�[�t�F�[�X�̃|�C���^�A
		���ʂ̒��̐��l�͎Q�ƃJ�E���^�̐��A���̎��̓N���X���ł��B
	*/
	public String toString() {
		String result = Integer.toHexString(pIUnknown) + "(" + (addRef()-1) + ")" + getClass().getName();
		release();		// ��LaddRef()�ŃJ�E���g�A�b�v�����Q�ƃJ�E���^���P���炷�B
		return result;
	}

	/**
		ReleaseManager��Ԃ��܂��B
		�ȉ��̌`�ŁA���݂̎Q�ƃJ�E���^�Ǘ��N���X�����邱�Ƃ��ł��܂��B
		<pre>System.out.println(excel.getReleaseManager().toString());</pre>
     * @see       #release()
	 */
	public ReleaseManager getReleaseManager() {	return rm; }

	// �i�m�h
	private native int _addRef();
	private native boolean _release();
	private native int _queryInterface(GUID IID) throws JComException;
}

