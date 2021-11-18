package jp.ne.so_net.ga2.no_ji.jcom.excel8;
import java.lang.*;
import jp.ne.so_net.ga2.no_ji.jcom.*;

public class ExcelWorksheets extends IDispatch {

	public ExcelWorksheets(IDispatch jcom) { super(jcom); }

	public ExcelApplication Application() throws JComException { return new ExcelApplication((IDispatch)get("Application")); }

	public int Creator() throws JComException { return ((Integer)get("Creator")).intValue(); }
//	long Worksheets::GetCreator()
//	{
//		long result;
//		InvokeHelper(0x95, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
//		return result;
//	}

//	LPDISPATCH Worksheets::GetParent()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x96, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}


	// ����m�F�ς�
	public ExcelWorksheet Add() throws JComException {
		return new ExcelWorksheet((IDispatch)method("Add",null));
	}
/* �g���Ȃ�
	// Before,After�͏ȗ����ɂ� null ���w�肵�Ă��������B
	// Type�ɂ�XlSheetType�N���X�̒萔���w�肵�܂��B����l�� XlSheetType.xlWorksheet�B
	public ExcelWorksheet Add(ExcelWorksheet Before, ExcelWorksheet After, int Count, int Type) throws JComException {
		Object[] arglist = new Object[4];
		arglist[0] = (Before==null)?(new IDispatch()):Before;
		arglist[1] = (After==null)?(new IDispatch()):After;
		arglist[2] = new Integer(Count);
		arglist[3] = new Integer(Type);
		return new ExcelWorksheet((IDispatch)method("Add",arglist));
	}
*/
//	LPDISPATCH Worksheets::Add(const VARIANT& Before, const VARIANT& After, const VARIANT& Count, const VARIANT& Type)
//	{
//		LPDISPATCH result;
//		static BYTE parms[] =
//			VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT;
//		InvokeHelper(0xb5, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms,
//			&Before, &After, &Count, &Type);
//		return result;
//	}


	public void Copy(ExcelWorksheet before, ExcelWorksheet after) throws JComException {
		method("Copy", new Object[] { before, after });
	}
//	void Worksheets::Copy(const VARIANT& Before, const VARIANT& After)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT VTS_VARIANT;
//		InvokeHelper(0x227, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 &Before, &After);
//	}
//	

	public int Count() throws JComException { return ((Integer)get("Count")).intValue(); }
//	long Worksheets::GetCount()
//	{
//		long result;
//		InvokeHelper(0x76, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
//		return result;
//	}

	public void Delete() throws JComException { method("Delete",null); }
//	void Worksheets::Delete()
//	{
//		InvokeHelper(0x75, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
//	}

//	void Worksheets::FillAcrossSheets(LPDISPATCH Range, long Type)
//	{
//		static BYTE parms[] =
//			VTS_DISPATCH VTS_I4;
//		InvokeHelper(0x1d5, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 Range, Type);
//	}

	// �w�肵���ԍ��̃��[�N�V�[�g��Ԃ��܂��B1�I���W���ł��B
	// ����m�F�ς�
	public ExcelWorksheet Item(int index) throws JComException {
		Object[] arglist = new Object[1];
		arglist[0] = new Integer(index);
		return new ExcelWorksheet((IDispatch)get("Item", arglist));
	}
//	LPDISPATCH Worksheets::GetItem(const VARIANT& Index)
//	{
//		LPDISPATCH result;
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0xaa, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, parms,
//			&Index);
//		return result;
//	}

//	void Worksheets::Move(const VARIANT& Before, const VARIANT& After)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT VTS_VARIANT;
//		InvokeHelper(0x27d, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 &Before, &After);
//	}
//	

	/**
		2000.11.27
		queryInterface()�̌Ăяo�����@��ύX�BIID��GUID�ł͂Ȃ��AIEnumVARIANT�̃��m���Q��
	*/
	// ����m�F�ς�
	public IEnumVARIANT NewEnum() throws JComException {
		IUnknown iUnknown = (IUnknown)get("_NewEnum");
//		Object a = iUnknown.queryInterface("jp.ne.so_net.ga2.no_ji.jcom.IEnumVARIANT", GUID.IID_IEnumVARIANT);
		Object a = iUnknown.queryInterface(IEnumVARIANT.class, IEnumVARIANT.IID);
		return (IEnumVARIANT)a;
	}
//	LPUNKNOWN Worksheets::Get_NewEnum()
//	{
//		LPUNKNOWN result;
//		InvokeHelper(0xfffffffc, DISPATCH_PROPERTYGET, VT_UNKNOWN, (void*)&result, NULL);
//		return result;
//	}


	// ����m�F�ς�
	public void PrintOut() throws JComException {
		method("PrintOut", null);
	}
//	void Worksheets::PrintOut(const VARIANT& From, const VARIANT& To, const VARIANT& Copies, const VARIANT& Preview, const VARIANT& ActivePrinter, const VARIANT& PrintToFile, const VARIANT& Collate)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT;
//		InvokeHelper(0x389, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 &From, &To, &Copies, &Preview, &ActivePrinter, &PrintToFile, &Collate);
//	}
//	
//	void Worksheets::PrintPreview(const VARIANT& EnableChanges)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0x119, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 &EnableChanges);
//	}
//	
//	void Worksheets::Select(const VARIANT& Replace)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0xeb, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 &Replace);
//	}
//	
//	LPDISPATCH Worksheets::GetHPageBreaks()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x58a, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}
//	
//	LPDISPATCH Worksheets::GetVPageBreaks()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x58b, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}
//	

	public boolean Visible() throws JComException { return ((Boolean)get("Visible")).booleanValue(); }
//	VARIANT Worksheets::GetVisible()
//	{
//		VARIANT result;
//		InvokeHelper(0x22e, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
//		return result;
//	}

	public void Visible(boolean v) throws JComException { put("Visible", new Boolean(v)); }
//	void Worksheets::SetVisible(const VARIANT& newValue)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0x22e, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 &newValue);
//	}

//	LPDISPATCH Worksheets::Get_Default(const VARIANT& Index)
//	{
//		LPDISPATCH result;
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, parms,
//			&Index);
//		return result;
//	}

	/**
		>javac jp/ne/so_net/ga2/no_ji/jcom/excel8/ExcelWorksheets.java
		>java jp/ne/so_net/ga2/no_ji/jcom/excel8/ExcelWorksheets		

		478698(1)jp.ne.so_net.ga2.no_ji.jcom.IDispatch
		478804(1)jp.ne.so_net.ga2.no_ji.jcom.IDispatch
		478970(1)jp.ne.so_net.ga2.no_ji.jcom.IDispatch
	*/
	public static void main(String[] args) throws JComException, java.io.IOException {
		ReleaseManager rm = new ReleaseManager();
		try {
			ExcelApplication excel = new ExcelApplication(rm);
			excel.Visible(true);
			ExcelWorkbooks xlBooks = excel.Workbooks();
			ExcelWorkbook xlBook = xlBooks.Add();	// �V�����u�b�N���쐬
			ExcelWorksheets xlSheets = xlBook.Worksheets();
			IEnumVARIANT enum = xlSheets.NewEnum();
			Object a = enum.next();
			do {
				System.out.println(""+a);
				a = enum.next();
			} while(a!=null);
			System.in.read();
			excel.Quit();
		} finally { rm.release(); }
	}

}
