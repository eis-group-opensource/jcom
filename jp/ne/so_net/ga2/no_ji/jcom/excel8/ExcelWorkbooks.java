// 修正履歴
// #01 2000-08-04 A.Watanabe         Openメソッドの戻り値を実装の間違い

package jp.ne.so_net.ga2.no_ji.jcom.excel8;
import jp.ne.so_net.ga2.no_ji.jcom.*;

public class ExcelWorkbooks extends IDispatch {

	ExcelWorkbooks(IDispatch disp) { super(disp); }

	public ExcelApplication Application() throws JComException { return new ExcelApplication((IDispatch)get("Application")); }
//	LPDISPATCH Workbooks::GetApplication()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x94, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}


	public int Creator() throws JComException { return ((Integer)get("Creator")).intValue(); }
//	long Workbooks::GetCreator()
//	{
//		long result;
//		InvokeHelper(0x95, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
//		return result;
//	}
//
//	LPDISPATCH Workbooks::GetParent()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x96, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}
//
     
	public static final int xlWBATExcel4MacroSheet     = 3;
	public static final int xlWBATExcel4IntlMacroSheet = 4;
	public static final int xlWBATChart                = -4109;
	public static final int xlWBATWorksheet            = -4167;
	public ExcelWorkbook Add(int Template) throws JComException {
		Object[] arglist = new Object[1];
		arglist[0] = new Integer(Template);
		return new ExcelWorkbook((IDispatch)method("Add", arglist));
	}
	public ExcelWorkbook Add() throws JComException {
		return new ExcelWorkbook((IDispatch)method("Add", null));
	}
//	LPDISPATCH Workbooks::Add(const VARIANT& Template)
//	{
//		LPDISPATCH result;
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0xb5, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms,
//			&Template);
//		return result;
//	}

	public void Close() throws JComException { method("Close", null); }
//	void Workbooks::Close()
//	{
//		InvokeHelper(0x115, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
//	}

	public int Count() throws JComException { return ((Integer)get("Count")).intValue(); }
//	long Workbooks::GetCount()
//	{
//		long result;
//		InvokeHelper(0x76, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
//		return result;
//	}


	public ExcelWorkbook Item(int index) throws JComException {
		Object[] arglist = new Object[1];
		arglist[0] = new Integer(index);
		return new ExcelWorkbook((IDispatch)get("Item",arglist));
	}
//	LPDISPATCH Workbooks::GetItem(const VARIANT& Index)
//	{
//		LPDISPATCH result;
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0xaa, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, parms,
//			&Index);
//		return result;
//	}


//	LPUNKNOWN Workbooks::Get_NewEnum()
//	{
//		LPUNKNOWN result;
//		InvokeHelper(0xfffffffc, DISPATCH_PROPERTYGET, VT_UNKNOWN, (void*)&result, NULL);
//		return result;
//	}


	//#01 begin
	/**
		引数はフルパスでないと駄目なのかな？
		相対パスはうまく行かない。パスなしだとMyDocumentsになる？
	*/
	public ExcelWorkbook Open(String Filename) throws JComException { Object args[] = new Object[1]; args[0] = Filename; return new ExcelWorkbook((IDispatch)method("Open", args)); }
	//public ExcelWorksheet Open(String Filename) throws JComException { Object args[] = new Object[1]; args[0] = Filename; return new ExcelWorksheet((IDispatch)method("Open", args)); }//@@@@@
	//#01 end

//	LPDISPATCH Workbooks::Open(LPCTSTR Filename, const VARIANT& UpdateLinks, const VARIANT& ReadOnly, const VARIANT& Format, const VARIANT& Password, const VARIANT& WriteResPassword, const VARIANT& IgnoreReadOnlyRecommended, const VARIANT& Origin, 
//			const VARIANT& Delimiter, const VARIANT& Editable, const VARIANT& Notify, const VARIANT& Converter, const VARIANT& AddToMru)
//	{
//		LPDISPATCH result;
//		static BYTE parms[] =
//			VTS_BSTR VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT;
//		InvokeHelper(0x2aa, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms,
//			Filename, &UpdateLinks, &ReadOnly, &Format, &Password, &WriteResPassword, &IgnoreReadOnlyRecommended, &Origin, &Delimiter, &Editable, &Notify, &Converter, &AddToMru);
//		return result;
//	}
//
//	void Workbooks::OpenText(LPCTSTR Filename, const VARIANT& Origin, const VARIANT& StartRow, const VARIANT& DataType, long TextQualifier, const VARIANT& ConsecutiveDelimiter, const VARIANT& Tab, const VARIANT& Semicolon, const VARIANT& Comma, 
//			const VARIANT& Space, const VARIANT& Other, const VARIANT& OtherChar, const VARIANT& FieldInfo, const VARIANT& TextVisualLayout)
//	{
//		static BYTE parms[] =
//			VTS_BSTR VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_I4 VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT;
//		InvokeHelper(0x2ab, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 Filename, &Origin, &StartRow, &DataType, TextQualifier, &ConsecutiveDelimiter, &Tab, &Semicolon, &Comma, &Space, &Other, &OtherChar, &FieldInfo, &TextVisualLayout);
//	}
//
//	LPDISPATCH Workbooks::Get_Default(const VARIANT& Index)
//	{
//		LPDISPATCH result;
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, parms,
//			&Index);
//		return result;
//	}

	public static void main(String[] args) throws JComException, java.io.IOException {
		ReleaseManager rm = new ReleaseManager();
		try {
			ExcelApplication excel = new ExcelApplication(rm);
			excel.Visible(true);
			ExcelWorkbooks xlBooks = excel.Workbooks();
			xlBooks.Open("a.xls");
			ExcelApplication xlApp = xlBooks.Application();
			System.out.println("Version="+xlApp.Version());
			System.in.read();
			excel.Quit();
		} finally { rm.release(); }
	}

}
