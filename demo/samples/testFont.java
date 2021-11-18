import jp.ne.so_net.ga2.no_ji.jcom.excel8.*;
import jp.ne.so_net.ga2.no_ji.jcom.*;
import java.io.File;
import java.util.Date;

/* sample of ExcelFont */
class testFont {
	public static void main(String[] args) throws Exception {
		ReleaseManager rm = new ReleaseManager();
		try {
			System.out.println("EXCEL ...");
			ExcelApplication excel = new ExcelApplication(rm);
			excel.Visible(true);
			// books - book - sheets - sheet - range
			ExcelWorkbooks xlBooks = excel.Workbooks();
			ExcelWorkbook xlBook = xlBooks.Add();	// create new book
			ExcelWorksheets xlSheets = xlBook.Worksheets();
			ExcelWorksheet xlSheet = xlSheets.Item(1);
			ExcelRange xlRange = xlSheet.Cells();

			// BOLD, Italic, Underline, etc...
			xlRange.Item(1,1).Value("BOLD");
			xlRange.Item(1,1).Font().Bold(true);
			xlRange.Item(1,2).Value("Italic");
			xlRange.Item(1,2).Font().Italic(true);
			xlRange.Item(1,3).Value("Underline");
			xlRange.Item(1,3).Font().Underline(true);
			xlRange.Item(1,4).Value("Strikethrough");
			xlRange.Item(1,4).Font().Strikethrough(true);
			xlRange.Item(1,5).Value("Subscript");
			xlRange.Item(1,5).Font().Subscript(true);
			xlRange.Item(1,6).Value("Superscript");
			xlRange.Item(1,6).Font().Superscript(true);
			xlRange.Item(1,7).Value("Shadow");
			xlRange.Item(1,7).Font().Shadow(true);

			// Colors
			xlRange.Item(2,1).Value("Color:red");
			xlRange.Item(2,1).Font().Color(0x0000FF);
			xlRange.Item(2,2).Value("Color:green");
			xlRange.Item(2,2).Font().Color(0x00FF00);
			xlRange.Item(2,3).Value("Color:blue");
			xlRange.Item(2,3).Font().Color(0xFF0000);

			// Font name
			xlRange.Item(3,1).Value("Arial");
			xlRange.Item(3,1).Font().Name("Arial");
			xlRange.Item(3,2).Value("Century");
			xlRange.Item(3,2).Font().Name("Century");
			xlRange.Item(3,3).Value("Roman");
			xlRange.Item(3,3).Font().Name("Roman");
			xlRange.Item(3,4).Value("Symbol");
			xlRange.Item(3,4).Font().Name("Symbol");
			xlRange.Item(3,5).Value("Wingdings");
			xlRange.Item(3,5).Font().Name("Wingdings");

			// Font size
			xlRange.Item(4,1).Value("size=8");
			xlRange.Item(4,1).Font().Size(8);
			xlRange.Item(4,2).Value("size=10");
			xlRange.Item(4,2).Font().Size(10);
			xlRange.Item(4,3).Value("size=24");
			xlRange.Item(4,3).Font().Size(24);

			System.out.println("Range(2,3).FontStyle="+xlRange.Item(2,3).Font().FontStyle());
			System.out.println("Range(2,3).Name="+xlRange.Item(2,3).Font().Name());

			System.out.println("hit [Enter] key to exit.");
			System.in.read();

			xlBook.Close(false,null,false);
			excel.Quit();

			System.out.println("thanks !");
		}
		catch(Exception e) { e.printStackTrace(); }
		finally { rm.release(); }
	}
}
/*
	javac -classpath ./jcom.jar testFont.java
	java -classpath ./jcom.jar;%CLASSPATH% testFont
*/
