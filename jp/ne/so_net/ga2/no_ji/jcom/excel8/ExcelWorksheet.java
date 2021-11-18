// 修正履歴
// #01 2000-08-04 A.Watanabe         main()を廃止


package jp.ne.so_net.ga2.no_ji.jcom.excel8;

import jp.ne.so_net.ga2.no_ji.jcom.*;

public class ExcelWorksheet extends IDispatch {

	public ExcelWorksheet(IDispatch jcom) { super(jcom); }

	public ExcelApplication Application() throws JComException { return new ExcelApplication((IDispatch)get("Application")); }
//	LPDISPATCH _Worksheet::GetApplication()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x94, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}

	public int Creator() throws JComException { return ((Integer)get("Creator")).intValue(); }
//	long _Worksheet::GetCreator()
//	{
//		long result;
//		InvokeHelper(0x95, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
//		return result;
//	}

//	LPDISPATCH _Worksheet::GetParent()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x96, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}

	public void Activate() throws JComException { method("Activate", null); }
//	void _Worksheet::Activate()
//	{
//		InvokeHelper(0x130, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
//	}

	public void Copy(ExcelWorksheet before, ExcelWorksheet after) throws JComException {
		method("Copy", new Object[] { before, after });
	}
//	void _Worksheet::Copy(const VARIANT& Before, const VARIANT& After)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT VTS_VARIANT;
//		InvokeHelper(0x227, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 &Before, &After);
//	}

	public void Delete() throws JComException { method("Delete",null); }
//	void _Worksheet::Delete()
//	{
//		InvokeHelper(0x75, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
//	}


	public String CodeName() throws JComException { return (String)get("CodeName"); }
//	CString _Worksheet::GetCodeName()
//	{
//		CString result;
//		InvokeHelper(0x55d, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
//		return result;
//	}
//
//	CString _Worksheet::Get_CodeName()
//	{
//		CString result;
//		InvokeHelper(0x80010000, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
//		return result;
//	}
//
//	void _Worksheet::Set_CodeName(LPCTSTR lpszNewValue)
//	{
//		static BYTE parms[] =
//			VTS_BSTR;
//		InvokeHelper(0x80010000, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 lpszNewValue);
//	}

	public int Index() throws JComException { return ((Integer)get("Index")).intValue(); }
//	long _Worksheet::GetIndex()
//	{
//		long result;
//		InvokeHelper(0x1e6, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
//		return result;
//	}

//	void _Worksheet::Move(const VARIANT& Before, const VARIANT& After)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT VTS_VARIANT;
//		InvokeHelper(0x27d, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 &Before, &After);
//	}

	public String Name() throws JComException { return (String)get("Name"); }
//	CString _Worksheet::GetName()
//	{
//		CString result;
//		InvokeHelper(0x6e, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
//		return result;
//	}

	public void Name(String newValue) throws JComException { put("Name", newValue); }
//	void _Worksheet::SetName(LPCTSTR lpszNewValue)
//	{
//		static BYTE parms[] =
//			VTS_BSTR;
//		InvokeHelper(0x6e, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 lpszNewValue);
//	}

	public ExcelWorksheet Next() throws JComException { return new ExcelWorksheet((IDispatch)get("Next")); }
//	LPDISPATCH _Worksheet::GetNext()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x1f6, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}

//	LPDISPATCH _Worksheet::GetPageSetup()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x3e6, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}

	public ExcelWorksheet Previous() throws JComException { return new ExcelWorksheet((IDispatch)get("Previous")); }
//	LPDISPATCH _Worksheet::GetPrevious()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x1f7, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}

	// 動作確認済み
	public void PrintOut() throws JComException {
		method("PrintOut", null);
	}
	public void PrintOut(int From, int To, int Copies, boolean Preview)
						throws JComException {
		Object[] arglist = new Object[4];
		arglist[0] = new Integer(From);
		arglist[1] = new Integer(To);
		arglist[2] = new Integer(Copies);
		arglist[3] = new Boolean(Preview);
		method("PrintOut", arglist);
	}
/* 動作しない
	public void PrintOut(int From, int To, int Copies, boolean Preview,
						String ActivePrinter, boolean PrintToFile, boolean Collate)
						throws JComException {
		Object[] arglist = new Object[7];
		arglist[0] = new Integer(From);
		arglist[1] = new Integer(To);
		arglist[2] = new Integer(Copies);
		arglist[3] = new Boolean(Preview);
		arglist[4] = ActivePrinter;
		arglist[5] = new Boolean(PrintToFile);
		arglist[6] = new Boolean(Collate);
		method("PrintOut", arglist);
	}
*/
//	void _Worksheet::PrintOut(const VARIANT& From, const VARIANT& To, const VARIANT& Copies, const VARIANT& Preview, const VARIANT& ActivePrinter, const VARIANT& PrintToFile, const VARIANT& Collate)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT;
//		InvokeHelper(0x389, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 &From, &To, &Copies, &Preview, &ActivePrinter, &PrintToFile, &Collate);
//	}
//
//	void _Worksheet::PrintPreview(const VARIANT& EnableChanges)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0x119, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 &EnableChanges);
//	}
//
//	void _Worksheet::Protect(const VARIANT& Password, const VARIANT& DrawingObjects, const VARIANT& Contents, const VARIANT& Scenarios, const VARIANT& UserInterfaceOnly)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT;
//		InvokeHelper(0x11a, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 &Password, &DrawingObjects, &Contents, &Scenarios, &UserInterfaceOnly);
//	}
//
//	BOOL _Worksheet::GetProtectContents()
//	{
//		BOOL result;
//		InvokeHelper(0x124, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
//		return result;
//	}
//
//	BOOL _Worksheet::GetProtectDrawingObjects()
//	{
//		BOOL result;
//		InvokeHelper(0x125, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
//		return result;
//	}
//
//	BOOL _Worksheet::GetProtectionMode()
//	{
//		BOOL result;
//		InvokeHelper(0x487, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
//		return result;
//	}
//
//	BOOL _Worksheet::GetProtectScenarios()
//	{
//		BOOL result;
//		InvokeHelper(0x126, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
//		return result;
//	}

	// FileFormatは XlFileFormatクラスの定数を指定してください。
	// 引数の数を減らしている。動作未確認
	// AccessModeに使用する定数 XlSaveAsAccessMode.xlNoChange (既定値) xlShared  xlExclusive 
	public void SaveAs(String Filename, int FileFormat, String Password, String WriteResPassword,
					boolean ReadOnlyRecommended, boolean CreateBackup, boolean AddToMru) throws JComException {
		Object[] arglist = new Object[7];
		arglist[0] = Filename;
		arglist[1] = new Integer(FileFormat);
		arglist[2] = Password;
		arglist[3] = WriteResPassword;
		arglist[4] = new Boolean(ReadOnlyRecommended);
		arglist[5] = new Boolean(CreateBackup);
		arglist[6] = new Boolean(AddToMru);
		method("SaveAs", arglist);
	}
//	void _Worksheet::SaveAs(LPCTSTR Filename, const VARIANT& FileFormat, 
//						const VARIANT& Password, const VARIANT& WriteResPassword,
//						const VARIANT& ReadOnlyRecommended, const VARIANT& CreateBackup,
//						const VARIANT& AddToMru, const VARIANT& TextCodepage
//			const VARIANT& TextVisualLayout)
//	{
//		static BYTE parms[] =
//			VTS_BSTR VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT;
//		InvokeHelper(0x11c, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 Filename, &FileFormat, &Password, &WriteResPassword, &ReadOnlyRecommended, &CreateBackup, &AddToMru, &TextCodepage, &TextVisualLayout);
//	}

//	void _Worksheet::Select(const VARIANT& Replace)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0xeb, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 &Replace);
//	}
//
//	void _Worksheet::Unprotect(const VARIANT& Password)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0x11d, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 &Password);
//	}
//
	public boolean Visible() throws JComException { return ((Boolean)get("Visible")).booleanValue(); }
//	long _Worksheet::GetVisible()
//	{
//		long result;
//		InvokeHelper(0x22e, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
//		return result;
//	}
//
	public void Visible(boolean v) throws JComException { put("Visible", new Boolean(v)); }
//	void _Worksheet::SetVisible(long nNewValue)
//	{
//		static BYTE parms[] =
//			VTS_I4;
//		InvokeHelper(0x22e, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 nNewValue);
//	}
//
//	LPDISPATCH _Worksheet::GetShapes()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x561, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}
//
//	BOOL _Worksheet::GetTransitionExpEval()
//	{
//		BOOL result;
//		InvokeHelper(0x191, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
//		return result;
//	}
//
//	void _Worksheet::SetTransitionExpEval(BOOL bNewValue)
//	{
//		static BYTE parms[] =
//			VTS_BOOL;
//		InvokeHelper(0x191, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 bNewValue);
//	}
//
//	BOOL _Worksheet::GetAutoFilterMode()
//	{
//		BOOL result;
//		InvokeHelper(0x318, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
//		return result;
//	}
//
//	void _Worksheet::SetAutoFilterMode(BOOL bNewValue)
//	{
//		static BYTE parms[] =
//			VTS_BOOL;
//		InvokeHelper(0x318, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 bNewValue);
//	}
//
//	void _Worksheet::SetBackgroundPicture(LPCTSTR Filename)
//	{
//		static BYTE parms[] =
//			VTS_BSTR;
//		InvokeHelper(0x4a4, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 Filename);
//	}

	public void Calculate() throws JComException { method("Calculate", null); }
//	void _Worksheet::Calculate()
//	{
//		InvokeHelper(0x117, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
//	}

	public boolean EnableCalculation() throws JComException { return ((Boolean)get("EnableCalculation")).booleanValue(); }
//	BOOL _Worksheet::GetEnableCalculation()
//	{
//		BOOL result;
//		InvokeHelper(0x590, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
//		return result;
//	}

	public void EnableCalculation(boolean newValue) throws JComException { put("EnableCalculation",new Boolean(newValue)); }
//	void _Worksheet::SetEnableCalculation(BOOL bNewValue)
//	{
//		static BYTE parms[] =
//			VTS_BOOL;
//		InvokeHelper(0x590, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 bNewValue);
//	}

	// すべてのRangeオブジェクトを返す。指定したセルは、Range.Item(RowIndex,ColumnIndex)等を使うこと
	public ExcelRange Cells() throws JComException { return new ExcelRange((IDispatch)get("Cells")); }
//	LPDISPATCH _Worksheet::GetCells()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0xee, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}

//	LPDISPATCH _Worksheet::ChartObjects(const VARIANT& Index)
//	{
//		LPDISPATCH result;
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0x424, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms,
//			&Index);
//		return result;
//	}
//
//	void _Worksheet::CheckSpelling(const VARIANT& CustomDictionary, const VARIANT& IgnoreUppercase, const VARIANT& AlwaysSuggest, const VARIANT& IgnoreInitialAlefHamza, const VARIANT& IgnoreFinalYaa, const VARIANT& SpellScript)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT;
//		InvokeHelper(0x1f9, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 &CustomDictionary, &IgnoreUppercase, &AlwaysSuggest, &IgnoreInitialAlefHamza, &IgnoreFinalYaa, &SpellScript);
//	}
//
//	LPDISPATCH _Worksheet::GetCircularReference()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x42d, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}
//
//	void _Worksheet::ClearArrows()
//	{
//		InvokeHelper(0x3ca, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
//	}

	public ExcelRange Columns() throws JComException { return new ExcelRange((IDispatch)get("Columns")); }
//	LPDISPATCH _Worksheet::GetColumns()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0xf1, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}

//	long _Worksheet::GetConsolidationFunction()
//	{
//		long result;
//		InvokeHelper(0x315, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
//		return result;
//	}
//
//	VARIANT _Worksheet::GetConsolidationOptions()
//	{
//		VARIANT result;
//		InvokeHelper(0x316, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
//		return result;
//	}
//
//	VARIANT _Worksheet::GetConsolidationSources()
//	{
//		VARIANT result;
//		InvokeHelper(0x317, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
//		return result;
//	}
//
//	BOOL _Worksheet::GetEnableAutoFilter()
//	{
//		BOOL result;
//		InvokeHelper(0x484, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
//		return result;
//	}
//
//	void _Worksheet::SetEnableAutoFilter(BOOL bNewValue)
//	{
//		static BYTE parms[] =
//			VTS_BOOL;
//		InvokeHelper(0x484, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 bNewValue);
//	}
//
//	long _Worksheet::GetEnableSelection()
//	{
//		long result;
//		InvokeHelper(0x591, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
//		return result;
//	}
//
//	void _Worksheet::SetEnableSelection(long nNewValue)
//	{
//		static BYTE parms[] =
//			VTS_I4;
//		InvokeHelper(0x591, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 nNewValue);
//	}
//
//	BOOL _Worksheet::GetEnableOutlining()
//	{
//		BOOL result;
//		InvokeHelper(0x485, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
//		return result;
//	}
//
//	void _Worksheet::SetEnableOutlining(BOOL bNewValue)
//	{
//		static BYTE parms[] =
//			VTS_BOOL;
//		InvokeHelper(0x485, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 bNewValue);
//	}
//
//	BOOL _Worksheet::GetEnablePivotTable()
//	{
//		BOOL result;
//		InvokeHelper(0x486, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
//		return result;
//	}
//
//	void _Worksheet::SetEnablePivotTable(BOOL bNewValue)
//	{
//		static BYTE parms[] =
//			VTS_BOOL;
//		InvokeHelper(0x486, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 bNewValue);
//	}
//
//	VARIANT _Worksheet::Evaluate(const VARIANT& Name)
//	{
//		VARIANT result;
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0x1, DISPATCH_METHOD, VT_VARIANT, (void*)&result, parms,
//			&Name);
//		return result;
//	}
//
//	VARIANT _Worksheet::_Evaluate(const VARIANT& Name)
//	{
//		VARIANT result;
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0xfffffffb, DISPATCH_METHOD, VT_VARIANT, (void*)&result, parms,
//			&Name);
//		return result;
//	}
//
//	BOOL _Worksheet::GetFilterMode()
//	{
//		BOOL result;
//		InvokeHelper(0x320, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
//		return result;
//	}
//
//	void _Worksheet::ResetAllPageBreaks()
//	{
//		InvokeHelper(0x592, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
//	}

//	LPDISPATCH _Worksheet::GetNames()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x1ba, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}
//
//	LPDISPATCH _Worksheet::OLEObjects(const VARIANT& Index)
//	{
//		LPDISPATCH result;
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0x31f, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms,
//			&Index);
//		return result;
//	}
//
//	LPDISPATCH _Worksheet::GetOutline()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x66, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}
//
//	void _Worksheet::Paste(const VARIANT& Destination, const VARIANT& Link)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT VTS_VARIANT;
//		InvokeHelper(0xd3, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 &Destination, &Link);
//	}
//
//	void _Worksheet::PasteSpecial(const VARIANT& Format, const VARIANT& Link, const VARIANT& DisplayAsIcon, const VARIANT& IconFileName, const VARIANT& IconIndex, const VARIANT& IconLabel)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT;
//		InvokeHelper(0x403, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 &Format, &Link, &DisplayAsIcon, &IconFileName, &IconIndex, &IconLabel);
//	}
//
//	LPDISPATCH _Worksheet::PivotTables(const VARIANT& Index)
//	{
//		LPDISPATCH result;
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0x2b2, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms,
//			&Index);
//		return result;
//	}
//
//	LPDISPATCH _Worksheet::PivotTableWizard(const VARIANT& SourceType, const VARIANT& SourceData, const VARIANT& TableDestination, const VARIANT& TableName, const VARIANT& RowGrand, const VARIANT& ColumnGrand, const VARIANT& SaveData, 
//			const VARIANT& HasAutoFormat, const VARIANT& AutoPage, const VARIANT& Reserved, const VARIANT& BackgroundQuery, const VARIANT& OptimizeCache, const VARIANT& PageFieldOrder, const VARIANT& PageFieldWrapCount, const VARIANT& ReadData, 
//			const VARIANT& Connection)
//	{
//		LPDISPATCH result;
//		static BYTE parms[] =
//			VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT;
//		InvokeHelper(0x2ac, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms,
//			&SourceType, &SourceData, &TableDestination, &TableName, &RowGrand, &ColumnGrand, &SaveData, &HasAutoFormat, &AutoPage, &Reserved, &BackgroundQuery, &OptimizeCache, &PageFieldOrder, &PageFieldWrapCount, &ReadData, &Connection);
//		return result;
//	}

	// Range("A1")  Range("A1:D10")  Range("TestRange")名前付き範囲
	public ExcelRange Range(String Cell1, String Cell2) throws JComException {
		Object[] arglist = new Object[2];
		arglist[0] = Cell1;
		arglist[1] = Cell2;
		return new ExcelRange((IDispatch)get("Range", arglist));
	}
	public ExcelRange Range(String Cell1) throws JComException {
		Object[] arglist = new Object[1];
		arglist[0] = Cell1;
		return new ExcelRange((IDispatch)get("Range", arglist));
	}

//	LPDISPATCH _Worksheet::GetRange(const VARIANT& Cell1, const VARIANT& Cell2)
//	{
//		LPDISPATCH result;
//		static BYTE parms[] =
//			VTS_VARIANT VTS_VARIANT;
//		InvokeHelper(0xc5, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, parms,
//			&Cell1, &Cell2);
//		return result;
//	}
//
	public ExcelRange Rows() throws JComException { return new ExcelRange((IDispatch)get("Rows")); }
//	LPDISPATCH _Worksheet::GetRows()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x102, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}

//	LPDISPATCH _Worksheet::Scenarios(const VARIANT& Index)
//	{
//		LPDISPATCH result;
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0x38c, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms,
//			&Index);
//		return result;
//	}
//
//	CString _Worksheet::GetScrollArea()
//	{
//		CString result;
//		InvokeHelper(0x599, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
//		return result;
//	}
//
//	void _Worksheet::SetScrollArea(LPCTSTR lpszNewValue)
//	{
//		static BYTE parms[] =
//			VTS_BSTR;
//		InvokeHelper(0x599, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 lpszNewValue);
//	}
//
//	void _Worksheet::ShowAllData()
//	{
//		InvokeHelper(0x31a, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
//	}
//
//	void _Worksheet::ShowDataForm()
//	{
//		InvokeHelper(0x199, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
//	}

	public double StandardHeight() throws JComException { return ((Double)get("StandardHeight")).doubleValue(); }
//	double _Worksheet::GetStandardHeight()
//	{
//		double result;
//		InvokeHelper(0x197, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, NULL);
//		return result;
//	}
//
	public double StandardWidth() throws JComException { return ((Double)get("StandardWidth")).doubleValue(); }
//	double _Worksheet::GetStandardWidth()
//	{
//		double result;
//		InvokeHelper(0x198, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, NULL);
//		return result;
//	}

	public void StandardHeight(double newValue) throws JComException { put("StandardHeight", new Double(newValue)); }
//	void _Worksheet::SetStandardWidth(double newValue)
//	{
//		static BYTE parms[] =
//			VTS_R8;
//		InvokeHelper(0x198, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 newValue);
//	}
//
//	BOOL _Worksheet::GetTransitionFormEntry()
//	{
//		BOOL result;
//		InvokeHelper(0x192, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
//		return result;
//	}
//
//	void _Worksheet::SetTransitionFormEntry(BOOL bNewValue)
//	{
//		static BYTE parms[] =
//			VTS_BOOL;
//		InvokeHelper(0x192, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 bNewValue);
//	}

	public int Type() throws JComException { return ((Integer)get("Type")).intValue(); }
//	long _Worksheet::GetType()
//	{
//		long result;
//		InvokeHelper(0x6c, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
//		return result;
//	}

	public ExcelRange UsedRange() throws JComException { return new ExcelRange((IDispatch)get("UsedRange")); }
//	LPDISPATCH _Worksheet::GetUsedRange()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x19c, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}

//	LPDISPATCH _Worksheet::GetHPageBreaks()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x58a, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}
//
//	LPDISPATCH _Worksheet::GetVPageBreaks()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x58b, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}
//
//	LPDISPATCH _Worksheet::GetQueryTables()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x59a, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}
//
//	BOOL _Worksheet::GetDisplayPageBreaks()
//	{
//		BOOL result;
//		InvokeHelper(0x59b, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
//		return result;
//	}
//
//	void _Worksheet::SetDisplayPageBreaks(BOOL bNewValue)
//	{
//		static BYTE parms[] =
//			VTS_BOOL;
//		InvokeHelper(0x59b, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 bNewValue);
//	}
//
//	LPDISPATCH _Worksheet::GetComments()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x23f, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}
//
//	LPDISPATCH _Worksheet::GetHyperlinks()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x571, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}
//
//	void _Worksheet::ClearCircles()
//	{
//		InvokeHelper(0x59c, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
//	}
//
//	void _Worksheet::CircleInvalid()
//	{
//		InvokeHelper(0x59d, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
//	}
//
//	LPDISPATCH _Worksheet::GetAutoFilter()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x319, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}

}
