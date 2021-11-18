// 修正履歴
// #01 2000-08-04 A.Watanabe         mainを廃止

package jp.ne.so_net.ga2.no_ji.jcom.excel8;
import java.lang.*;
import jp.ne.so_net.ga2.no_ji.jcom.*;

public class ExcelWorkbook extends IDispatch {

	public ExcelWorkbook(IDispatch disp) { super(disp); }

	public ExcelApplication Application() throws JComException { return new ExcelApplication((IDispatch)get("Application")); }
//	LPDISPATCH _Workbook::GetApplication()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x94, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}


	public int Creator() throws JComException { return ((Integer)get("Creator")).intValue(); }
//	long _Workbook::GetCreator()
//	{
//		long result;
//		InvokeHelper(0x95, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
//		return result;
//	}
//
//	LPDISPATCH _Workbook::GetParent()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x96, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}
//
//	BOOL _Workbook::GetAcceptLabelsInFormulas()
//	{
//		BOOL result;
//		InvokeHelper(0x5a1, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
//		return result;
//	}
//
//	void _Workbook::SetAcceptLabelsInFormulas(BOOL bNewValue)
//	{
//		static BYTE parms[] =
//			VTS_BOOL;
//		InvokeHelper(0x5a1, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 bNewValue);
//	}
//
//	void _Workbook::Activate()
//	{
//		InvokeHelper(0x130, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
//	}
//
//	LPDISPATCH _Workbook::GetActiveChart()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0xb7, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}

	public ExcelWorksheet ActiveSheet() throws JComException { return new ExcelWorksheet((IDispatch)get("ActiveSheet")); }
//	LPDISPATCH _Workbook::GetActiveSheet()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x133, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}


//	long _Workbook::GetAutoUpdateFrequency()
//	{
//		long result;
//		InvokeHelper(0x5a2, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
//		return result;
//	}
//
//	void _Workbook::SetAutoUpdateFrequency(long nNewValue)
//	{
//		static BYTE parms[] =
//			VTS_I4;
//		InvokeHelper(0x5a2, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 nNewValue);
//	}
//
//	BOOL _Workbook::GetAutoUpdateSaveChanges()
//	{
//		BOOL result;
//		InvokeHelper(0x5a3, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
//		return result;
//	}
//
//	void _Workbook::SetAutoUpdateSaveChanges(BOOL bNewValue)
//	{
//		static BYTE parms[] =
//			VTS_BOOL;
//		InvokeHelper(0x5a3, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 bNewValue);
//	}
//
//	long _Workbook::GetChangeHistoryDuration()
//	{
//		long result;
//		InvokeHelper(0x5a4, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
//		return result;
//	}
//
//	void _Workbook::SetChangeHistoryDuration(long nNewValue)
//	{
//		static BYTE parms[] =
//			VTS_I4;
//		InvokeHelper(0x5a4, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 nNewValue);
//	}
//
//	LPDISPATCH _Workbook::GetBuiltinDocumentProperties()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x498, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}
//
//	void _Workbook::ChangeFileAccess(long Mode, const VARIANT& WritePassword, const VARIANT& Notify)
//	{
//		static BYTE parms[] =
//			VTS_I4 VTS_VARIANT VTS_VARIANT;
//		InvokeHelper(0x3dd, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 Mode, &WritePassword, &Notify);
//	}
//
//	void _Workbook::ChangeLink(LPCTSTR Name, LPCTSTR NewName, long Type)
//	{
//		static BYTE parms[] =
//			VTS_BSTR VTS_BSTR VTS_I4;
//		InvokeHelper(0x322, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 Name, NewName, Type);
//	}
//
//	LPDISPATCH _Workbook::GetCharts()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x79, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}

	public void Close(boolean SaveChanges, String Filename, boolean RouteWorkbook) throws JComException {
		Object[] arglist = new Object[3];
		arglist[0] = new Boolean(SaveChanges);
		arglist[1] = Filename;
		arglist[2] = new Boolean(RouteWorkbook);
		method("Close", arglist);
	}
//	void _Workbook::Close(const VARIANT& SaveChanges, const VARIANT& Filename, const VARIANT& RouteWorkbook)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT VTS_VARIANT VTS_VARIANT;
//		InvokeHelper(0x115, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 &SaveChanges, &Filename, &RouteWorkbook);
//	}
//
//	CString _Workbook::GetCodeName()
//	{
//		CString result;
//		InvokeHelper(0x55d, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
//		return result;
//	}
//
//	CString _Workbook::Get_CodeName()
//	{
//		CString result;
//		InvokeHelper(0x80010000, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
//		return result;
//	}
//
//	void _Workbook::Set_CodeName(LPCTSTR lpszNewValue)
//	{
//		static BYTE parms[] =
//			VTS_BSTR;
//		InvokeHelper(0x80010000, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 lpszNewValue);
//	}
//
//	VARIANT _Workbook::GetColors(const VARIANT& Index)
//	{
//		VARIANT result;
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0x11e, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, parms,
//			&Index);
//		return result;
//	}
//
//	void _Workbook::SetColors(const VARIANT& Index, const VARIANT& newValue)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT VTS_VARIANT;
//		InvokeHelper(0x11e, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 &Index, &newValue);
//	}
//
//	LPDISPATCH _Workbook::GetCommandBars()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x59f, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}
//
//	long _Workbook::GetConflictResolution()
//	{
//		long result;
//		InvokeHelper(0x497, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
//		return result;
//	}
//
//	void _Workbook::SetConflictResolution(long nNewValue)
//	{
//		static BYTE parms[] =
//			VTS_I4;
//		InvokeHelper(0x497, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 nNewValue);
//	}
//
//	LPDISPATCH _Workbook::GetContainer()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x4a6, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}
//
//	BOOL _Workbook::GetCreateBackup()
//	{
//		BOOL result;
//		InvokeHelper(0x11f, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
//		return result;
//	}
//
//	LPDISPATCH _Workbook::GetCustomDocumentProperties()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x499, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}
//
//	BOOL _Workbook::GetDate1904()
//	{
//		BOOL result;
//		InvokeHelper(0x193, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
//		return result;
//	}
//
//	void _Workbook::SetDate1904(BOOL bNewValue)
//	{
//		static BYTE parms[] =
//			VTS_BOOL;
//		InvokeHelper(0x193, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 bNewValue);
//	}
//
//	void _Workbook::DeleteNumberFormat(LPCTSTR NumberFormat)
//	{
//		static BYTE parms[] =
//			VTS_BSTR;
//		InvokeHelper(0x18d, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 NumberFormat);
//	}
//
//	long _Workbook::GetDisplayDrawingObjects()
//	{
//		long result;
//		InvokeHelper(0x194, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
//		return result;
//	}
//
//	void _Workbook::SetDisplayDrawingObjects(long nNewValue)
//	{
//		static BYTE parms[] =
//			VTS_I4;
//		InvokeHelper(0x194, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 nNewValue);
//	}
//
//	BOOL _Workbook::ExclusiveAccess()
//	{
//		BOOL result;
//		InvokeHelper(0x490, DISPATCH_METHOD, VT_BOOL, (void*)&result, NULL);
//		return result;
//	}
//
	public int FileFormat() throws JComException { return ((Integer)get("FileFormat")).intValue(); }
//	long _Workbook::GetFileFormat()
//	{
//		long result;
//		InvokeHelper(0x120, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
//		return result;
//	}

//	void _Workbook::ForwardMailer()
//	{
//		InvokeHelper(0x3cd, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
//	}
//
//	CString _Workbook::GetFullName()
//	{
//		CString result;
//		InvokeHelper(0x121, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
//		return result;
//	}
//
//	BOOL _Workbook::GetHasPassword()
//	{
//		BOOL result;
//		InvokeHelper(0x122, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
//		return result;
//	}
//
//	BOOL _Workbook::GetHasRoutingSlip()
//	{
//		BOOL result;
//		InvokeHelper(0x3b6, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
//		return result;
//	}
//
//	void _Workbook::SetHasRoutingSlip(BOOL bNewValue)
//	{
//		static BYTE parms[] =
//			VTS_BOOL;
//		InvokeHelper(0x3b6, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 bNewValue);
//	}
//
//	BOOL _Workbook::GetIsAddin()
//	{
//		BOOL result;
//		InvokeHelper(0x5a5, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
//		return result;
//	}
//
//	void _Workbook::SetIsAddin(BOOL bNewValue)
//	{
//		static BYTE parms[] =
//			VTS_BOOL;
//		InvokeHelper(0x5a5, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 bNewValue);
//	}
//
//	VARIANT _Workbook::LinkInfo(LPCTSTR Name, long LinkInfo, const VARIANT& Type, const VARIANT& EditionRef)
//	{
//		VARIANT result;
//		static BYTE parms[] =
//			VTS_BSTR VTS_I4 VTS_VARIANT VTS_VARIANT;
//		InvokeHelper(0x327, DISPATCH_METHOD, VT_VARIANT, (void*)&result, parms,
//			Name, LinkInfo, &Type, &EditionRef);
//		return result;
//	}
//
//	VARIANT _Workbook::LinkSources(const VARIANT& Type)
//	{
//		VARIANT result;
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0x328, DISPATCH_METHOD, VT_VARIANT, (void*)&result, parms,
//			&Type);
//		return result;
//	}
//
//	LPDISPATCH _Workbook::GetMailer()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x3d3, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}
//
//	void _Workbook::MergeWorkbook(const VARIANT& Filename)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0x5a6, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 &Filename);
//	}
//
//	BOOL _Workbook::GetMultiUserEditing()
//	{
//		BOOL result;
//		InvokeHelper(0x491, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
//		return result;
//	}
//
//	CString _Workbook::GetName()
//	{
//		CString result;
//		InvokeHelper(0x6e, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
//		return result;
//	}
//
//	LPDISPATCH _Workbook::GetNames()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x1ba, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}
//
//	LPDISPATCH _Workbook::NewWindow()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x118, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}
//
//	void _Workbook::OpenLinks(LPCTSTR Name, const VARIANT& ReadOnly, const VARIANT& Type)
//	{
//		static BYTE parms[] =
//			VTS_BSTR VTS_VARIANT VTS_VARIANT;
//		InvokeHelper(0x323, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 Name, &ReadOnly, &Type);
//	}
//
//	CString _Workbook::GetPath()
//	{
//		CString result;
//		InvokeHelper(0x123, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
//		return result;
//	}
//
//	BOOL _Workbook::GetPersonalViewListSettings()
//	{
//		BOOL result;
//		InvokeHelper(0x5a7, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
//		return result;
//	}
//
//	void _Workbook::SetPersonalViewListSettings(BOOL bNewValue)
//	{
//		static BYTE parms[] =
//			VTS_BOOL;
//		InvokeHelper(0x5a7, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 bNewValue);
//	}
//
//	BOOL _Workbook::GetPersonalViewPrintSettings()
//	{
//		BOOL result;
//		InvokeHelper(0x5a8, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
//		return result;
//	}
//
//	void _Workbook::SetPersonalViewPrintSettings(BOOL bNewValue)
//	{
//		static BYTE parms[] =
//			VTS_BOOL;
//		InvokeHelper(0x5a8, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 bNewValue);
//	}
//
//	LPDISPATCH _Workbook::PivotCaches()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x5a9, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}
//
//	void _Workbook::Post(const VARIANT& DestName)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0x48e, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 &DestName);
//	}
//
//	BOOL _Workbook::GetPrecisionAsDisplayed()
//	{
//		BOOL result;
//		InvokeHelper(0x195, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
//		return result;
//	}
//
//	void _Workbook::SetPrecisionAsDisplayed(BOOL bNewValue)
//	{
//		static BYTE parms[] =
//			VTS_BOOL;
//		InvokeHelper(0x195, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 bNewValue);
//	}

	// 動作確認済み
	public void PrintOut() throws JComException {
		method("PrintOut", null);
	}
//	void _Workbook::PrintOut(const VARIANT& From, const VARIANT& To, const VARIANT& Copies, const VARIANT& Preview, const VARIANT& ActivePrinter, const VARIANT& PrintToFile, const VARIANT& Collate)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT;
//		InvokeHelper(0x389, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 &From, &To, &Copies, &Preview, &ActivePrinter, &PrintToFile, &Collate);
//	}
//
//	void _Workbook::PrintPreview(const VARIANT& EnableChanges)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0x119, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 &EnableChanges);
//	}
//
//	void _Workbook::Protect(const VARIANT& Password, const VARIANT& Structure, const VARIANT& Windows)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT VTS_VARIANT VTS_VARIANT;
//		InvokeHelper(0x11a, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 &Password, &Structure, &Windows);
//	}
//
//	void _Workbook::ProtectSharing(const VARIANT& Filename, const VARIANT& Password, const VARIANT& WriteResPassword, const VARIANT& ReadOnlyRecommended, const VARIANT& CreateBackup, const VARIANT& SharingPassword)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT;
//		InvokeHelper(0x5aa, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 &Filename, &Password, &WriteResPassword, &ReadOnlyRecommended, &CreateBackup, &SharingPassword);
//	}
//
//	BOOL _Workbook::GetProtectStructure()
//	{
//		BOOL result;
//		InvokeHelper(0x24c, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
//		return result;
//	}
//
//	BOOL _Workbook::GetProtectWindows()
//	{
//		BOOL result;
//		InvokeHelper(0x127, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
//		return result;
//	}
//
//	BOOL _Workbook::GetReadOnly()
//	{
//		BOOL result;
//		InvokeHelper(0x128, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
//		return result;
//	}
//
//	BOOL _Workbook::GetReadOnlyRecommended()
//	{
//		BOOL result;
//		InvokeHelper(0x129, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
//		return result;
//	}
//
//	void _Workbook::RefreshAll()
//	{
//		InvokeHelper(0x5ac, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
//	}
//
//	void _Workbook::Reply()
//	{
//		InvokeHelper(0x3d1, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
//	}
//
//	void _Workbook::ReplyAll()
//	{
//		InvokeHelper(0x3d2, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
//	}
//
//	void _Workbook::RemoveUser(long Index)
//	{
//		static BYTE parms[] =
//			VTS_I4;
//		InvokeHelper(0x5ad, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 Index);
//	}
//
//	long _Workbook::GetRevisionNumber()
//	{
//		long result;
//		InvokeHelper(0x494, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
//		return result;
//	}
//
//	void _Workbook::Route()
//	{
//		InvokeHelper(0x3b2, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
//	}
//
//	BOOL _Workbook::GetRouted()
//	{
//		BOOL result;
//		InvokeHelper(0x3b7, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
//		return result;
//	}
//
//	LPDISPATCH _Workbook::GetRoutingSlip()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x3b5, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}
//
//	void _Workbook::RunAutoMacros(long Which)
//	{
//		static BYTE parms[] =
//			VTS_I4;
//		InvokeHelper(0x27a, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 Which);
//	}

	public void Save() throws JComException { method("Save",null); }
//	void _Workbook::Save()
//	{
//		InvokeHelper(0x11b, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
//	}






	public void SaveAs(String Filename) throws JComException {
		Object[] arglist = new Object[1];
		arglist[0] = Filename;
		method("SaveAs", arglist);
	}

/* 未完成
	public void SaveAs(String Filename, int FileFormat) throws JComException {
		Object[] arglist = new Object[2];
		arglist[0] = Filename;
		arglist[1] = new Integer(FileFormat);
		method("SaveAs", arglist);
	}
	// FileFormatは XlFileFormatクラスの定数を指定してください。xlWorkbookNormal???
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
*/
//	void _Workbook::SaveAs(const VARIANT& Filename, const VARIANT& FileFormat,
//						 const VARIANT& Password, const VARIANT& WriteResPassword,
//						 const VARIANT& ReadOnlyRecommended, const VARIANT& CreateBackup,
//						 long AccessMode, const VARIANT& ConflictResolution, 
//						 const VARIANT& AddToMru, const VARIANT& TextCodepage, const VARIANT& TextVisualLayout)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_I4 VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT;
//		InvokeHelper(0x11c, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 &Filename, &FileFormat, &Password, &WriteResPassword, &ReadOnlyRecommended, &CreateBackup, AccessMode, &ConflictResolution, &AddToMru, &TextCodepage, &TextVisualLayout);
//	}
//
//	void _Workbook::SaveCopyAs(const VARIANT& Filename)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0xaf, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 &Filename);
//	}
//
//	BOOL _Workbook::GetSaved()
//	{
//		BOOL result;
//		InvokeHelper(0x12a, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
//		return result;
//	}
//
//	void _Workbook::SetSaved(BOOL bNewValue)
//	{
//		static BYTE parms[] =
//			VTS_BOOL;
//		InvokeHelper(0x12a, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 bNewValue);
//	}
//
//	BOOL _Workbook::GetSaveLinkValues()
//	{
//		BOOL result;
//		InvokeHelper(0x196, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
//		return result;
//	}
//
//	void _Workbook::SetSaveLinkValues(BOOL bNewValue)
//	{
//		static BYTE parms[] =
//			VTS_BOOL;
//		InvokeHelper(0x196, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 bNewValue);
//	}
//
//	void _Workbook::SendMail(const VARIANT& Recipients, const VARIANT& Subject, const VARIANT& ReturnReceipt)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT VTS_VARIANT VTS_VARIANT;
//		InvokeHelper(0x3b3, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 &Recipients, &Subject, &ReturnReceipt);
//	}
//
//	void _Workbook::SendMailer(const VARIANT& FileFormat, long Priority)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT VTS_I4;
//		InvokeHelper(0x3d4, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 &FileFormat, Priority);
//	}
//
//	void _Workbook::SetLinkOnData(LPCTSTR Name, const VARIANT& Procedure)
//	{
//		static BYTE parms[] =
//			VTS_BSTR VTS_VARIANT;
//		InvokeHelper(0x329, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 Name, &Procedure);
//	}
//
//	LPDISPATCH _Workbook::GetSheets()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x1e5, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}
//
//	BOOL _Workbook::GetShowConflictHistory()
//	{
//		BOOL result;
//		InvokeHelper(0x493, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
//		return result;
//	}
//
//	void _Workbook::SetShowConflictHistory(BOOL bNewValue)
//	{
//		static BYTE parms[] =
//			VTS_BOOL;
//		InvokeHelper(0x493, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 bNewValue);
//	}
//
//	LPDISPATCH _Workbook::GetStyles()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x1ed, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}
//
//	void _Workbook::Unprotect(const VARIANT& Password)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0x11d, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 &Password);
//	}
//
//	void _Workbook::UnprotectSharing(const VARIANT& SharingPassword)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0x5af, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 &SharingPassword);
//	}
//
//	void _Workbook::UpdateFromFile()
//	{
//		InvokeHelper(0x3e3, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
//	}
//
//	void _Workbook::UpdateLink(const VARIANT& Name, const VARIANT& Type)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT VTS_VARIANT;
//		InvokeHelper(0x324, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 &Name, &Type);
//	}
//
//	BOOL _Workbook::GetUpdateRemoteReferences()
//	{
//		BOOL result;
//		InvokeHelper(0x19b, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
//		return result;
//	}
//
//	void _Workbook::SetUpdateRemoteReferences(BOOL bNewValue)
//	{
//		static BYTE parms[] =
//			VTS_BOOL;
//		InvokeHelper(0x19b, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 bNewValue);
//	}
//
//	VARIANT _Workbook::GetUserStatus()
//	{
//		VARIANT result;
//		InvokeHelper(0x495, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
//		return result;
//	}
//
//	LPDISPATCH _Workbook::GetCustomViews()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x5b0, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}
//
//	LPDISPATCH _Workbook::GetWindows()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x1ae, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}
//

	public ExcelWorksheets Worksheets() throws JComException { return new ExcelWorksheets((IDispatch)get("Worksheets")); }
//	LPDISPATCH _Workbook::GetWorksheets()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x1ee, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}
//
//	BOOL _Workbook::GetWriteReserved()
//	{
//		BOOL result;
//		InvokeHelper(0x12b, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
//		return result;
//	}
//
//	CString _Workbook::GetWriteReservedBy()
//	{
//		CString result;
//		InvokeHelper(0x12c, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
//		return result;
//	}
//
//	LPDISPATCH _Workbook::GetExcel4IntlMacroSheets()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x245, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}
//
//	LPDISPATCH _Workbook::GetExcel4MacroSheets()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x243, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}
//
//	BOOL _Workbook::GetTemplateRemoveExtData()
//	{
//		BOOL result;
//		InvokeHelper(0x5b1, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
//		return result;
//	}
//
//	void _Workbook::SetTemplateRemoveExtData(BOOL bNewValue)
//	{
//		static BYTE parms[] =
//			VTS_BOOL;
//		InvokeHelper(0x5b1, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 bNewValue);
//	}
//
//	void _Workbook::HighlightChangesOptions(const VARIANT& When, const VARIANT& Who, const VARIANT& Where)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT VTS_VARIANT VTS_VARIANT;
//		InvokeHelper(0x5b2, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 &When, &Who, &Where);
//	}
//
//	BOOL _Workbook::GetHighlightChangesOnScreen()
//	{
//		BOOL result;
//		InvokeHelper(0x5b5, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
//		return result;
//	}
//
//	void _Workbook::SetHighlightChangesOnScreen(BOOL bNewValue)
//	{
//		static BYTE parms[] =
//			VTS_BOOL;
//		InvokeHelper(0x5b5, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 bNewValue);
//	}
//
//	BOOL _Workbook::GetKeepChangeHistory()
//	{
//		BOOL result;
//		InvokeHelper(0x5b6, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
//		return result;
//	}
//
//	void _Workbook::SetKeepChangeHistory(BOOL bNewValue)
//	{
//		static BYTE parms[] =
//			VTS_BOOL;
//		InvokeHelper(0x5b6, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 bNewValue);
//	}
//
//	BOOL _Workbook::GetListChangesOnNewSheet()
//	{
//		BOOL result;
//		InvokeHelper(0x5b7, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
//		return result;
//	}
//
//	void _Workbook::SetListChangesOnNewSheet(BOOL bNewValue)
//	{
//		static BYTE parms[] =
//			VTS_BOOL;
//		InvokeHelper(0x5b7, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 bNewValue);
//	}
//
//	void _Workbook::PurgeChangeHistoryNow(long Days, const VARIANT& SharingPassword)
//	{
//		static BYTE parms[] =
//			VTS_I4 VTS_VARIANT;
//		InvokeHelper(0x5b8, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 Days, &SharingPassword);
//	}
//
//	void _Workbook::AcceptAllChanges(const VARIANT& When, const VARIANT& Who, const VARIANT& Where)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT VTS_VARIANT VTS_VARIANT;
//		InvokeHelper(0x5ba, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 &When, &Who, &Where);
//	}
//
//	void _Workbook::RejectAllChanges(const VARIANT& When, const VARIANT& Who, const VARIANT& Where)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT VTS_VARIANT VTS_VARIANT;
//		InvokeHelper(0x5bb, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 &When, &Who, &Where);
//	}
//
//	void _Workbook::ResetColors()
//	{
//		InvokeHelper(0x5bc, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
//	}
//
//	LPDISPATCH _Workbook::GetVBProject()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x5bd, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}
//
//	void _Workbook::FollowHyperlink(LPCTSTR Address, const VARIANT& SubAddress, const VARIANT& NewWindow, const VARIANT& AddHistory, const VARIANT& ExtraInfo, const VARIANT& Method, const VARIANT& HeaderInfo)
//	{
//		static BYTE parms[] =
//			VTS_BSTR VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT;
//		InvokeHelper(0x5be, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 Address, &SubAddress, &NewWindow, &AddHistory, &ExtraInfo, &Method, &HeaderInfo);
//	}
//
//	void _Workbook::AddToFavorites()
//	{
//		InvokeHelper(0x5c4, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
//	}
//
//	BOOL _Workbook::GetIsInplace()
//	{
//		BOOL result;
//		InvokeHelper(0x6f4, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
//		return result;
//	}
}

