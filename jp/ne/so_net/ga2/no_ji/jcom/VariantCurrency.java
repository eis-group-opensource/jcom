package jp.ne.so_net.ga2.no_ji.jcom;

import java.text.NumberFormat;

/**
 * VARIANT�^�̒ʉ݌^���`���܂��B
 * VARIANT�ł�CY/CURRENCY��LONGLONG�Ƃ��Ď�������Ă��܂����A
 * Java�ł͒ʏ�ʉ݂�double���Ă��܂��B�����I�ɒʉ݂�double��
 * ��ʂ��邽�߂ɁA���̂悤�Ɏ������܂����B
 *
 * @see     IDispatch
 * @see     JComException
	@author Yoshinori Watanabe(�n�� �`��)
	@version 2.00, 2000/06/25
	Copyright(C) Yoshinori Watanabe 1999-2000. All Rights Reserved.
 */
public class VariantCurrency {
	double value;
	static NumberFormat price = NumberFormat.getCurrencyInstance();

    /**
     * �w�肳�ꂽ���z��VariantCurrency���쐬���܂��B
     * @param     value ���z
     */
	public VariantCurrency(double value) { this.value = value; }

    /**
     * VariantCurrency���쐬���܂��B���z�͂O�ŏ���������܂��B
     */
	public VariantCurrency() { this(0.0); }

    /**
     * �w�肵�����z��ݒ肵�܂��B
     * @param	value	���z
     */
	public void set(double value) { this.value = value; }

	/**
     * ���z���擾���܂��B
     * @return	���z
     */
	public double get() { return value; }

    /**
     * ���z�𕶎���ɕϊ����܂��B
     * ������
     * <code>NumberFormat.getCurrencyInstance()</code>
     * �ɏ]���܂��B
     * @return	������
     */
	public String toString() {
		return price.format(value);
	}
}
