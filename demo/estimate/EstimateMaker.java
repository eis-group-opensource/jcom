import jp.ne.so_net.ga2.no_ji.jcom.excel8.*;
import jp.ne.so_net.ga2.no_ji.jcom.*;
import java.io.File;
import java.util.Date;

/*
	Excel用ラッパを使った、JComのサンプルプログラム
*/
public class EstimateMaker {

	public boolean print_enable = false;	// 印刷するかどうか
	// 設定項目
	public String company     = "";				// 会社名
	public String section1    = "";				// 所属1
	public String section2    = "";				// 所属2
	public String custmer     = "";				// 顧客名
	public String validperiod = "";				// 見積有効期限
	public String createdate  = "";				// 作成日
	public String estimatedNo = "";				// 見積書No.
	public String charge      = "";				// 担当者
	public String[] itemname = new String[15];	// 品名x15個
	public String[] itemtype = new String[15];	// 型番x15個
	public int[] itemprice   = new int[15];		// 単価x15個
	public int[] itemcount   = new int[15];		// 数量x15個
	public String[] itemmemo = new String[15];	// 備考x15個

	public boolean makeEstimate(String fname) {
		ReleaseManager rm = new ReleaseManager();
		try {
			
			System.out.println("EXCELを起動中...");
			// すでに立ち上がっていると、新しいウィンドウで開く。
			ExcelApplication excel = new ExcelApplication(rm);
			excel.Visible(true);

			ExcelWorkbooks xlBooks = excel.Workbooks();
			ExcelWorkbook xlBook = xlBooks.Open(fname);
			ExcelWorksheet xlSheet = excel.ActiveSheet();
			ExcelRange xlRange = xlSheet.Cells();
			
			// 設定項目をセルに代入
			System.out.println("設定項目をセルに代入");
			xlRange.Item(4,2).Value(company);
			xlRange.Item(5,2).Value(section1);
			xlRange.Item(6,2).Value(section2);
			xlRange.Item(8,2).Value(custmer);
			xlRange.Item(14,4).Value(validperiod);
			xlRange.Item(3,8).Value(createdate);
			xlRange.Item(5,8).Value(estimatedNo);
			xlRange.Item(7,8).Value(charge);
			for(int i=0; i<15; i++) {
				xlRange.Item(22+i,3).Value(itemname[i]);
				xlRange.Item(22+i,4).Value(itemtype[i]);
				if(itemcount[i]!=0) {
					xlRange.Item(22+i,5).Value(itemprice[i]);
					xlRange.Item(22+i,6).Value(itemcount[i]);
				}
				else {
					xlRange.Item(22+i,5).Value("");
					xlRange.Item(22+i,6).Value("");
				}
				xlRange.Item(22+i,8).Value(itemmemo[i]);
			}

			// プリンタに出力する場合はコメントをはずしてください。
			// デフォルトのプリンタに出力されます。
			if(print_enable) {
				System.out.println("プリンタに印刷します。");
				xlSheet.PrintOut();
			}

			// ファイルに保存する場合はコメントを外してください。
			// ディレクトリを指定しない場合は、(My Documents)に保存されます。
			System.out.println("ファイルに保存します。");
			xlBook.Save();

			xlBook.Close(false,null,false);
			excel.Quit();
		}
		catch(Exception e) {
			e.printStackTrace();
			return false;	// 失敗
		}
		finally { rm.release(); }
		return true;
	}

	public static void main(String[] args) {
		EstimateMaker est = new EstimateMaker();
		est.company     = "まっくろそふと";			// 会社名
		est.section1    = "あぷり開発部門";			// 所属1
		est.section2    = "まくせるグループ";		// 所属2 えくせるだろ！
		est.custmer     = "びる げえつ 様";			// 顧客名
		est.validperiod = "2000/09/05";				// 見積有効期限
		est.createdate  = "2000/08/06";				// 作成日
		est.estimatedNo = "PL1234-56-7890";			// 見積書No.
		est.charge      = "渡辺 義則";				// 担当者
		for(int i=0; i<15; i++) {
			est.itemname[i]  = "-";		// 品名x15
			est.itemtype[i]  = "";		// 型番x15
			est.itemprice[i] = 0;		// 単価x15
			est.itemcount[i] = 0;		// 数量x15
			est.itemmemo[i]  = "";		// 備考x15
		}
		// 項目1
		est.itemname[0] = "関数ツリーVer14.40";
		est.itemtype[0] = "FT-1440S";
		est.itemprice[0] = 1500;
		est.itemcount[0] = 1;
		est.itemmemo[0]  = "ソース付";
		// 項目2
		est.itemname[1] = "JCom 2.00";
		est.itemtype[1] = "JC-200";
		est.itemprice[1] = 800;
		est.itemcount[1] = 4;
		est.itemmemo[1]  = "";
		// 項目3
		est.itemname[2] = "SYLBIS α版";
		est.itemtype[2] = "SY-1A";
		est.itemprice[2] = 3000;
		est.itemcount[2] = 1;
		est.itemmemo[2]  = "DirectX対応";
		// 値を設定＆印刷.
		est.print_enable = false;
		// 作業用のファイルを作成。見積書Noと同じにする。
		try {
			String workfile = ".\\"+est.estimatedNo+".xls";
			System.out.println("ファイルをコピー "+workfile);
			FileCopy.copy(".\\estimate.xls", workfile);
			// 見積書を作成。Excelは独自の環境を持つので、ファイル名を絶対パスにして渡す
			System.out.println("見積書作成");
			boolean rc = est.makeEstimate((new File(workfile)).getCanonicalPath());
			if(rc)
				System.out.println("成功しました");
			else
				System.out.println("失敗しました(;_;)");
		}
		catch(Exception e) {
			e.printStackTrace();
			System.out.println("失敗しました(T_T)");
		}
	}
}
