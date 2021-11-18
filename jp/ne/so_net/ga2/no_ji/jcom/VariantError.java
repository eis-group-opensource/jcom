package jp.ne.so_net.ga2.no_ji.jcom;


/**
 * VARIANT型のERROR型を定義します。
 * scode は HRESULT の値です。意味はCOMのヘルプや WINERROR.H などを参照してください。
 * @see     IDispatch
 * @see     JComException
	@author Yoshinori Watanabe(渡辺 義則)
	@version 2.24, 2004/05/31
	Copyright(C) Yoshinori Watanabe 1999-2004. All Rights Reserved.
 */
public class VariantError {
	int scode;

    /**
     * 指定されたエラーコードでVariantErrorを作成します。
     * @param     scode エラーコード
     */
	public VariantError(int scode) { this.scode = scode; }

    /**
     * VariantErrorを作成します。エラーコードは 0 で初期化されます。
     */
	public VariantError() { this(0); }

    /**
     * 指定したエラーコードを設定します。
     * @param     scode エラーコード
     */
	public void set(int scode) { this.scode = scode; }

	/**
     * エラーコードを取得します。
     * @return	エラーコード
     */
	public int get() { return scode; }

    /**
     * エラーコードを文字列に変換します。
     * @return	文字列
     */
	public String toString() {
		return Integer.toHexString(scode).toUpperCase();
	}
}
