package jp.ne.so_net.ga2.no_ji.jcom;


/**
 * VARIANT�^��ERROR�^���`���܂��B
 * scode �� HRESULT �̒l�ł��B�Ӗ���COM�̃w���v�� WINERROR.H �Ȃǂ��Q�Ƃ��Ă��������B
 * @see     IDispatch
 * @see     JComException
	@author Yoshinori Watanabe(�n�� �`��)
	@version 2.24, 2004/05/31
	Copyright(C) Yoshinori Watanabe 1999-2004. All Rights Reserved.
 */
public class VariantError {
	int scode;

    /**
     * �w�肳�ꂽ�G���[�R�[�h��VariantError���쐬���܂��B
     * @param     scode �G���[�R�[�h
     */
	public VariantError(int scode) { this.scode = scode; }

    /**
     * VariantError���쐬���܂��B�G���[�R�[�h�� 0 �ŏ���������܂��B
     */
	public VariantError() { this(0); }

    /**
     * �w�肵���G���[�R�[�h��ݒ肵�܂��B
     * @param     scode �G���[�R�[�h
     */
	public void set(int scode) { this.scode = scode; }

	/**
     * �G���[�R�[�h���擾���܂��B
     * @return	�G���[�R�[�h
     */
	public int get() { return scode; }

    /**
     * �G���[�R�[�h�𕶎���ɕϊ����܂��B
     * @return	������
     */
	public String toString() {
		return Integer.toHexString(scode).toUpperCase();
	}
}
