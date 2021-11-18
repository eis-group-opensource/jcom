package jp.ne.so_net.ga2.no_ji.jcom;
import java.lang.reflect.*;

/**
 * IPersist�C���^�[�t�F�[�X���������߂̃N���X
 	���̃N���X��CLSID���擾���邽�߂����ɂ���܂��B
 	�ȉ��̕��@�ŁA���̃C���^�[�t�F�[�X���T�|�[�g���Ă���
 	�b�n�l�I�u�W�F�N�g�ɑ΂���CLSID�A��������ProgID��
 	�擾���邱�Ƃ��\�ł��B�������A���ׂẴI�u�W�F�N�g��
 	���̃C���^�[�t�F�[�X���T�|�[�g���Ă���킯�ł͂���܂���B
	Excel�̏ꍇ�AExcel.Application�ł͎g���܂���B
	Excel.Sheet, Excel.Chart ���o�[�W�����t���̌`��ProgID��
	�Ԃ��܂��B("Excel.Chart.8"�̌`��)
	���ƁA���[�h��Word.Document�Ȃǂ��Ԃ��܂��B
	�ȉ���ProgID���擾���邽�߂̃T���v���ł��B
	<PRE>
 *   public static String getProgID(IUnknown unknown) {
 *       try {
 *           IPersist persist = (IPersist)unknown.queryInterface(IPersist.class, IPersist.IID);
 *           if(persist==null) return null;
 *           GUID clsid = persist.getClassID();
 *           return Com.getProgIDFromCLSID(clsid);
 *       }
 *       catch(JComException e) { e.printStackTrace(); }
 *       return null;
 *   }</PRE>
 *
 * @see     IUnknown
 * @see     JComException
 * @see     ReleaseManager
	@author Yoshinori Watanabe(�n�� �`��)
	@version 2.21, 2000/11/27
	Copyright(C) Yoshinori Watanabe 1999-2000. All Rights Reserved.
 */
public class IPersist extends IUnknown {
    /**
		IID_IPersist �ł��B0000010c-0000-0000-C000-000000000046
		@see       GUID
	*/
	public static GUID IID = new GUID( 0x0000010C, 0x0000, 0x0000, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46 );

    /**
     * IPersist���쐬���܂��B
     * �����Ŏw�肳�ꂽIPersist�|�C���^��COM���쐬���܂��B
     * �ʏ�A�ʂ�COM�̃��\�b�h����Ԃ��ꂽIPersist�|�C���^�ɑ΂��āA
     * �g�p���܂��B
     * @param     rm             �Q�ƃJ�E���^�Ǘ��N���X
     * @param     pIPersist     IPersist�C���^�[�t�F�[�X�̃A�h���X
     * @see       ReleaseManager
	 */
	public IPersist(ReleaseManager rm, int pIPersist) {
		super(rm, pIPersist);
	}

	/**
		CLSID��Ԃ��܂��B
	*/
	public synchronized GUID getClassID() throws JComException {
		return _getClassID();
	}

	// release()��super�̂łn�j�D

	// �i�m�h
	private native GUID _getClassID() throws JComException;
}

