package jp.ne.so_net.ga2.no_ji.jcom;
/**
	JCom ���ŗ�O�����������Ƃ��ɃX���[����܂��B
	��O�̓��e�ɂ��Ă�getMessage()���s���ĉ������B
	
	
	
	
	@see Exception
	@see IUnknown
	@see IDispatch
	@author Yoshinori Watanabe(�n�� �`��)
	@version 2.00, 2000/06/25
	Copyright(C) Yoshinori Watanabe 1999-2000. All Rights Reserved.
*/
public class JComException extends Exception {
    /**
     * ��O���쐬���܂��B
		@see IUnknown
		@see IDispatch
	 */
	JComException(String msg) {
		super(msg);
	}
}
