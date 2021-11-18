package jp.ne.so_net.ga2.no_ji.jcom.excel8;
/*
保存するときのアクセスモード
Workbook.SaveAs()等で使用
*/
class XlSaveAsAccessMode {
	static final int xlNoChange  = 1;	// 既定値
	static final int xlShared    = 2;
	static final int xlExclusive = 3;
}
