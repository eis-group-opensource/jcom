package jp.ne.so_net.ga2.no_ji.jcom;

/**
 *  �R���N�V�����I�u�W�F�N�g���������߂̃N���X
 * IEnumVARIANT�C���^�[�t�F�[�X�ɂ�Clone()�Ƃ������\�b�h������܂����A
 * ����ɂ͑Ή����Ă��܂���BNext(),Reset(),Skip()�ɂ̂ݑΉ����Ă��܂��B
 * �܂��ANext()�ɂ͗��p�ړI�ɍ��킹�A�Q��ނ̊֐���p�ӂ��Ă��܂��B
 *
	@author Yoshinori Watanabe(�n�� �`��)
	@version 2.00, 2000/06/25
	Copyright(C) Yoshinori Watanabe 1999-2000. All Rights Reserved.
 * @see     IUnknown
 */
public class IEnumVARIANT extends IUnknown {
    /**
		IID_IEnumVARIANT �ł��B
		@see       GUID
	*/
	public static GUID IID = new GUID( 0x00020404, 0x0000, 0x0000, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46 );

    /**
     * IEnumVARIANT���쐬���܂��B
	 * ����pIEnumVARIANT��IUnknown.queryInterface()���g����
	 	�擾�����l���w�肵�܂��B
     * @param     rm             �Q�ƃJ�E���^�Ǘ��N���X
     	@param	pIEnumVARIANT	pIEnumVARIANT�C���^�[�t�F�[�X
     * @see       ReleaseManager
	 */
	public IEnumVARIANT(ReleaseManager rm, int pIEnumVARIANT) {
		super(rm, pIEnumVARIANT);
	}

	/**
		�P���̃I�u�W�F�N�g�����o���܂��B
		���̃I�u�W�F�N�g���Ȃ��ꍇ��null��Ԃ��܂��B
	*/
	public synchronized Object next() throws JComException {
		Object ret = _next();
		if(rm!=null && (ret instanceof IUnknown)) {
			rm.add((IUnknown)ret);
		}
		return ret;
	}

	/**
		�w�肵�����������̃I�u�W�F�N�g�����o���܂��B
		�c�肪���Ȃ��ꍇ�́A�w�肵�����ȉ��ɂȂ�ꍇ������܂��B
		�z��̗v�f���ɒ��ӂ��ĉ������B
		celt�͂P�ȏ�̐����w�肵�ĉ������B
		@param	celt	�擾����I�u�W�F�N�g�̐�(1�`)
	*/
	public synchronized Object[] next(int celt) throws JComException {
		Object[] ary = _next(celt);
		if(rm!=null) {
			for(int i=0; i<ary.length; i++) {
				if(ary[i] instanceof IUnknown) {
					rm.add((IUnknown)ary[i]);
				}
			}
		}
		return ary;
	}

	/**
		�ŏ������蒼���܂��B
		�J�[�\�����ŏ��Ɉړ����܂��B
	*/
	public synchronized void reset() throws JComException {
		_reset();
	}

	/**
		�w�肵���������I�u�W�F�N�g���X�L�b�v�����܂��B
		celt�͂P�ȏ�̐����w�肵�ĉ������B
		@param	celt	�X�L�b�v�����鐔(1�`)
	*/
	public synchronized void skip(int celt) throws JComException {
		_skip(celt);
	}

	// JNI
	private native Object _next() throws JComException;
	private native Object[] _next(int celt) throws JComException;
	private native void _reset() throws JComException;
	private native void _skip(int celt) throws JComException;
}
