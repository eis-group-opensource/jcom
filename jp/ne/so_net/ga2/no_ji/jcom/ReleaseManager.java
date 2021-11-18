package jp.ne.so_net.ga2.no_ji.jcom;
import java.util.*;

/**
 * ReleaseManager �Q�ƃJ�E���^�Ǘ��N���X�B
 * ������Ȃ���΂Ȃ�Ȃ��I�u�W�F�N�g���Ǘ����܂��B
 *
 *	��r�I�ȒP�ȃv���O�����ł͈ȉ��̌`�ŗ��p���ĉ������B
 *   <pre>
 *   // ��r�I�Z���̃v���O����
 *   ReleaseManager rm = new ReleaseManager();
 *   try {
 *       IDispatch foo = new IDispatch(rm ,progid);
 *       // ...
 *   } catch(JComException e) {
 *       e.printStackTrace();
 *   } finally {
 *       rm.release();
 *   }
 *   </pre>
 * �܂��A�T�[�o�[�A�v���P�[�V������A���G�Œ����ԓ��삷��A�v���P�[�V����
 * �ł́AReleaseManager�̐�������release()���s���܂łɁA�I�u�W�F�N�g���������
 * ����΂Ȃ�Ȃ��ꍇ������܂��B�I�u�W�F�N�g��������Ȃ��ƁA����������������
 * ���܂�����ł��B���̏ꍇ�́A�K���ȏ����P�ʂ�push()��pop()�ň͂ނ��Ƃɂ��A
 * ���̒��Ŋm�ۂ��ꂽ�I�u�W�F�N�g��������܂��B
 * push()��pop()�͕K���΂ɂȂ�悤�ɂ��Ă��������B
 * ���̑΂� push() pop() push() pop()�Ƃ����ӂ��ɁA������ĂԂ��Ƃ��ł��܂��B
 * �܂��Apush() push() pop() pop()�Ƃ����ӂ��ɁA�l�X�g���邱�Ƃ��ł��܂��B
 * �ȉ��̗�ł́Afor���̒P�ʂŃI�u�W�F�N�g��������Ă��܂��B
 *      <pre>
 *        // ��r�I�����̒����v���O����
 *        ReleaseManager rm = new ReleaseManager();
 *        try {
 *            IDispatch foo = new IDispatch(rm ,progid);
 *            //  ...
 *            for(int i=0; i&lt;files.length; i++) {
 *                rm.push();
 *                // ...
 *                rm.pop(); // for���̒��Ő������ꂽ�I�u�W�F�N�g�����
 *            }
 *        } catch(JComException e) {
 *            e.printStackTrace();
 *        } finally {
 *            rm.release();
 *        }
 *      </pre>
 * �����̂b�n�l�𓯎��Ɉ����Ƃ��A���̐����Ɖ���̃^�C�~���O���قȂ�ꍇ������܂��B
 * �Ⴆ�΁A�P�̂c�a���b�n�l�Ƃ��Ĉ����A�����̂d�w�b�d�k���܂��b�n�l�Ƃ��Ĉ����ꍇ�ł��B
 * �c�a�͍ŏ��ɂP�񐶐�����̂ɑ΂��A�d�w�b�d�k�͕����񐶐����邱�ƂɂȂ邩��ł��B
 * ���̏ꍇ�͕�����<code> ReleaseManager </code>�𐶐����A���ꂼ��Ɋ��蓖�Ă邱�Ƃɂ��A
 * �I�u�W�F�N�g�̉���ɂ��ăL���ׂ̍���������s�����Ƃ��ł��܂��B
 * �ʂ̂b�n�l�ɑ΂��āA����<code> ReleaseManager </code>���g�����Ƃ��A�قȂ�
 * <code> ReleaseManager </code>���g�����Ƃ��ł��܂��B�܂��A��������ӏ���
 * �ʂ̃u���b�N�i���\�b�h�A�X���b�h���j�ɂ��邱�Ƃ��ł��܂��B
 *      <pre>
 *        // �����̈قȂ镡���̂b�n�l�������v���O�����i��P�j
 *        ReleaseManager rmDb = new ReleaseManager();
 *        ReleaseManager rmExcel = new ReleaseManager();
 *        try {
 *            IDispatch comDB = new IDispatch(rmDb ,"foo.DB");
 *            IDispatch comExcel = new IDispatch(rmExcel ,"Excel.Application");
 *            //  ...
 *            for(int i=0; i&lt;table.length; i++) {
 *                rmExcel.push();
 *                // ...
 *                rmExcel.pop();  //�d�w�b�d�k�I�u�W�F�N�g�̂݉��
 *            }
 *        } catch(JComException e) {
 *            e.printStackTrace();
 *        } finally {
 *            rmExcel.release();
 *            rmDb.release();
 *        }
 *      </pre>
 *      <pre>
 *        // �����̈قȂ镡���̂b�n�l�������v���O����(��Q)
 *        ReleaseManager rmDB = new ReleaseManager();
 *        try {
 *            IDispatch comDB = new IDispatch(rmDB ,"foo.DB");
 *            rmDB.push();
 *            ReleaseManager rmExcel = new ReleaseManager();
 *            try {
 *                IDispatch comExcel = new IDispatch(rmExcel ,"Excel.Application");
 *                //  ...
 *                for(int i=0; i&lt;table.length; i++) {
 *                    rmExcel.push();
 *                    rmDB.push();
 *                    // ...
 *                    rmDB.pop();       // �c�a�I�u�W�F�N�g�̂݉��
 *                    rmExcel.pop();   // �d�w�b�d�k�I�u�W�F�N�g�����
 *                }
 *            } catch(JComException e) {
 *                e.printStackTrace();
 *            } finally {
 *                rmExcel.release();
 *            }
 *            rmDB.pop();          //  �c�a�I�u�W�F�N�g�̂݉��
 *        } catch(JComException e) {
 *            e.printStackTrace();
 *        } finally {
 *            rmDB.release();
 *        }
 *      </pre>
	@author Yoshinori Watanabe(�n�� �`��)
	@version 2.00, 2000/06/25
	Copyright(C) Yoshinori Watanabe 1999-2000. All Rights Reserved.
 	@see     IDispatch
 	@see	IUnknown
 */
public class ReleaseManager {

	private Stack frames;
	private Stack curFrame;
	/**
		ReleaseManager���쐬���܂��B
		������Ȃ���΂Ȃ�Ȃ��b�n�l�I�u�W�F�N�g���Ǘ����܂��B
	*/
	public ReleaseManager() {
		frames = new Stack();
		curFrame = new Stack();
		frames.push(curFrame);
	}
	/**
		IUnknown �����݂̃X�^�b�N�ɒǉ����܂��B
	*/
	public void add(IUnknown jcom) {
		curFrame.push(jcom);
	}

	/**
		�V�����X�^�b�N�𐶐����܂��B
	*/
	public void push() {
		curFrame = new Stack();
		frames.push(curFrame);
	}
	/**
		���݂̃X�^�b�N���IUnknown��������܂��B
		���̌�A1�O�̃X�^�b�N�ɖ߂��܂��B
	*/
	public void pop() {
		if(curFrame==null) return;
		// ���݂̃t���[���ɗ��܂���JCom��release����
		while( ! curFrame.empty() ) {
			((IUnknown)curFrame.pop()).release();
		}
		frames.pop();
		// �t���[������O�ɖ߂�
		try {
			curFrame = (Stack)frames.peek();
		} catch(EmptyStackException e) {
			curFrame = null;
		}
	}
	/**
		���ׂẴX�^�b�N���IUnknown��������܂��B
	*/
	public void release() {
		while( ! frames.empty() ) {
			pop();
		}
	}
	/**
		���ׂẴX�^�b�N���IUnknown��������܂��B
		�i�������u�l�̓v���O�����I������<code> filnalize() </code>���ĂԂ��Ƃ�
		�ۏ؂��Ă��܂���B
		�ʏ�́Atry�`catch����JComException���L���b�`���A
		finally����release()�𖾎��I�ɌĂԂ悤�ɂ��Ă��������B
	*/
	public void finalize() {
		release();
	}

	/**
		�����ŕێ����Ă���IUnknown�N���X�A�܂��͂��ꂩ��p�������N���X�̃I�u�W�F�N�g
		���ȉ��̌`�ŕ\�����܂��B
		�P�U�i���̓C���^�[�t�F�[�X�̃|�C���^�A
		���ʂ̒��̐��l�͎Q�ƃJ�E���^�̐��A���̎��̓N���X���ł��B
		�C�ӂ̉ӏ��ł̃I�u�W�F�N�g�̏�Ԃ��X�i�b�v�V���b�g�I�Ɍ��邱�Ƃ��ł��܂��B
		<pre>
        {
        {
        4769e4(1)jp.ne.so_net.ga2.no_ji.jcom.excel8.ExcelApplication
        476eb8(1)jp.ne.so_net.ga2.no_ji.jcom.IDispatch
        477c98(1)jp.ne.so_net.ga2.no_ji.jcom.IDispatch
        477ed4(1)jp.ne.so_net.ga2.no_ji.jcom.IDispatch
        478004(1)jp.ne.so_net.ga2.no_ji.jcom.IUnknown
        478694(2)jp.ne.so_net.ga2.no_ji.jcom.IDispatch
        4788f4(2)jp.ne.so_net.ga2.no_ji.jcom.IDispatch
        478b30(2)jp.ne.so_net.ga2.no_ji.jcom.IDispatch
        478694(2)jp.ne.so_net.ga2.no_ji.jcom.IDispatch
        4788f4(2)jp.ne.so_net.ga2.no_ji.jcom.IDispatch
        478b30(2)jp.ne.so_net.ga2.no_ji.jcom.IDispatch
        }
        }</pre>
	*/
	public String toString() {
		String result = "{\n";
		for(int i=0; i<frames.size(); i++) {
			Stack s = (Stack)frames.elementAt(i);
			result += "{\n";
			for(int j=0; j<s.size(); j++) {
				IUnknown p = (IUnknown)s.elementAt(j);
				result += p.toString() + "\n";
			}
			result += "}\n";
		}
		result += "}\n";
		return result;
	}
}

