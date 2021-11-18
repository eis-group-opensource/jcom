package jp.ne.so_net.ga2.no_ji.jcom.excel8;
import java.lang.*;
import java.util.*;
import jp.ne.so_net.ga2.no_ji.jcom.*;

public class ExcelRange extends IDispatch {

	public ExcelRange(IDispatch disp) { super(disp); }

	public ExcelApplication Application() throws JComException { return new ExcelApplication((IDispatch)get("Application")); }
//	LPDISPATCH Range::GetApplication()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x94, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}


	public int Creator() throws JComException { return ((Integer)get("Creator")).intValue(); }
//	long Range::GetCreator()
//	{
//		long result;
//		InvokeHelper(0x95, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
//		return result;
//	}
//
//	LPDISPATCH Range::GetParent()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x96, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}
//
//	void Range::Activate()
//	{
//		InvokeHelper(0x130, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
//	}
//
//	VARIANT Range::GetAddIndent()
//	{
//		VARIANT result;
//		InvokeHelper(0x427, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
//		return result;
//	}
//
//	void Range::SetAddIndent(const VARIANT& newValue)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0x427, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 &newValue);
//	}
//
//	CString Range::GetAddress(const VARIANT& RowAbsolute, const VARIANT& ColumnAbsolute, long ReferenceStyle, const VARIANT& External, const VARIANT& RelativeTo)
//	{
//		CString result;
//		static BYTE parms[] =
//			VTS_VARIANT VTS_VARIANT VTS_I4 VTS_VARIANT VTS_VARIANT;
//		InvokeHelper(0xec, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, parms,
//			&RowAbsolute, &ColumnAbsolute, ReferenceStyle, &External, &RelativeTo);
//		return result;
//	}
//
//	CString Range::GetAddressLocal(const VARIANT& RowAbsolute, const VARIANT& ColumnAbsolute, long ReferenceStyle, const VARIANT& External, const VARIANT& RelativeTo)
//	{
//		CString result;
//		static BYTE parms[] =
//			VTS_VARIANT VTS_VARIANT VTS_I4 VTS_VARIANT VTS_VARIANT;
//		InvokeHelper(0x1b5, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, parms,
//			&RowAbsolute, &ColumnAbsolute, ReferenceStyle, &External, &RelativeTo);
//		return result;
//	}
//
//	void Range::AdvancedFilter(long Action, const VARIANT& CriteriaRange, const VARIANT& CopyToRange, const VARIANT& Unique)
//	{
//		static BYTE parms[] =
//			VTS_I4 VTS_VARIANT VTS_VARIANT VTS_VARIANT;
//		InvokeHelper(0x36c, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 Action, &CriteriaRange, &CopyToRange, &Unique);
//	}
//
//	void Range::ApplyNames(const VARIANT& Names, const VARIANT& IgnoreRelativeAbsolute, const VARIANT& UseRowColumnNames, const VARIANT& OmitColumn, const VARIANT& OmitRow, long Order, const VARIANT& AppendLast)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_I4 VTS_VARIANT;
//		InvokeHelper(0x1b9, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 &Names, &IgnoreRelativeAbsolute, &UseRowColumnNames, &OmitColumn, &OmitRow, Order, &AppendLast);
//	}
//
//	void Range::ApplyOutlineStyles()
//	{
//		InvokeHelper(0x1c0, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
//	}
//
//	LPDISPATCH Range::GetAreas()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x238, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}
//
//	CString Range::AutoComplete(LPCTSTR String)
//	{
//		CString result;
//		static BYTE parms[] =
//			VTS_BSTR;
//		InvokeHelper(0x4a1, DISPATCH_METHOD, VT_BSTR, (void*)&result, parms,
//			String);
//		return result;
//	}
//
//	void Range::AutoFill(LPDISPATCH Destination, long Type)
//	{
//		static BYTE parms[] =
//			VTS_DISPATCH VTS_I4;
//		InvokeHelper(0x1c1, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 Destination, Type);
//	}
//
//	void Range::AutoFilter(const VARIANT& Field, const VARIANT& Criteria1, long Operator, const VARIANT& Criteria2, const VARIANT& VisibleDropDown)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT VTS_VARIANT VTS_I4 VTS_VARIANT VTS_VARIANT;
//		InvokeHelper(0x319, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 &Field, &Criteria1, Operator, &Criteria2, &VisibleDropDown);
//	}
//
	public void AutoFit() throws JComException { method("AutoFit",null); }
//	void Range::AutoFit()
//	{
//		InvokeHelper(0xed, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
//	}

//	void Range::AutoFormat(long Format, const VARIANT& Number, const VARIANT& Font, const VARIANT& Alignment, const VARIANT& Border, const VARIANT& Pattern, const VARIANT& Width)
//	{
//		static BYTE parms[] =
//			VTS_I4 VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT;
//		InvokeHelper(0x72, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 Format, &Number, &Font, &Alignment, &Border, &Pattern, &Width);
//	}
//
//	void Range::AutoOutline()
//	{
//		InvokeHelper(0x40c, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
//	}
//
//	void Range::BorderAround(const VARIANT& LineStyle, long Weight, long ColorIndex, const VARIANT& Color)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT VTS_I4 VTS_I4 VTS_VARIANT;
//		InvokeHelper(0x42b, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 &LineStyle, Weight, ColorIndex, &Color);
//	}
//
//	LPDISPATCH Range::GetBorders()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x1b3, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}

	public void Calculate() throws JComException { method("Calculate", null); }
//	void Range::Calculate()
//	{
//		InvokeHelper(0x117, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
//	}

	public ExcelRange Cells() throws JComException { return new ExcelRange((IDispatch)get("Cells")); }
//	LPDISPATCH Range::GetCells()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0xee, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}
//
//	LPDISPATCH Range::GetCharacters(const VARIANT& Start, const VARIANT& Length)
//	{
//		LPDISPATCH result;
//		static BYTE parms[] =
//			VTS_VARIANT VTS_VARIANT;
//		InvokeHelper(0x25b, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, parms,
//			&Start, &Length);
//		return result;
//	}
//
//	void Range::CheckSpelling(const VARIANT& CustomDictionary, const VARIANT& IgnoreUppercase, const VARIANT& AlwaysSuggest, const VARIANT& IgnoreInitialAlefHamza, const VARIANT& IgnoreFinalYaa, const VARIANT& SpellScript)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT;
//		InvokeHelper(0x1f9, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 &CustomDictionary, &IgnoreUppercase, &AlwaysSuggest, &IgnoreInitialAlefHamza, &IgnoreFinalYaa, &SpellScript);
//	}

	public void Clear() throws JComException { method("Clear",null); }
//	void Range::Clear()
//	{
//		InvokeHelper(0x6f, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
//	}

//	void Range::ClearContents()
//	{
//		InvokeHelper(0x71, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
//	}
//
//	void Range::ClearFormats()
//	{
//		InvokeHelper(0x70, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
//	}
//
//	void Range::ClearNotes()
//	{
//		InvokeHelper(0xef, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
//	}
//
//	void Range::ClearOutline()
//	{
//		InvokeHelper(0x40d, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
//	}

	public int Column() throws JComException { return ((Integer)get("Column")).intValue(); }
//	long Range::GetColumn()
//	{
//		long result;
//		InvokeHelper(0xf0, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
//		return result;
//	}
//
//	LPDISPATCH Range::ColumnDifferences(const VARIANT& Comparison)
//	{
//		LPDISPATCH result;
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0x1fe, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms,
//			&Comparison);
//		return result;
//	}

	public ExcelRange Columns() throws JComException { return new ExcelRange((IDispatch)get("Columns")); }
//	LPDISPATCH Range::GetColumns()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0xf1, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}
//
//	VARIANT Range::GetColumnWidth()
//	{
//		VARIANT result;
//		InvokeHelper(0xf2, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
//		return result;
//	}
//
//	void Range::SetColumnWidth(const VARIANT& newValue)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0xf2, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 &newValue);
//	}
//
//	void Range::Consolidate(const VARIANT& Sources, const VARIANT& Function, const VARIANT& TopRow, const VARIANT& LeftColumn, const VARIANT& CreateLinks)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT;
//		InvokeHelper(0x1e2, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 &Sources, &Function, &TopRow, &LeftColumn, &CreateLinks);
//	}
//
	public void Copy(ExcelRange destination) throws JComException {
		method("Copy", new Object[] { destination } );
	}
//	void Range::Copy(const VARIANT& Destination)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0x227, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 &Destination);
//	}
//
//	long Range::CopyFromRecordset(LPUNKNOWN Data, const VARIANT& MaxRows, const VARIANT& MaxColumns)
//	{
//		long result;
//		static BYTE parms[] =
//			VTS_UNKNOWN VTS_VARIANT VTS_VARIANT;
//		InvokeHelper(0x480, DISPATCH_METHOD, VT_I4, (void*)&result, parms,
//			Data, &MaxRows, &MaxColumns);
//		return result;
//	}
//
//	void Range::CopyPicture(long Appearance, long Format)
//	{
//		static BYTE parms[] =
//			VTS_I4 VTS_I4;
//		InvokeHelper(0xd5, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 Appearance, Format);
//	}

	public int Count() throws JComException { return ((Integer)get("Count")).intValue(); }
//	long Range::GetCount()
//	{
//		long result;
//		InvokeHelper(0x76, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
//		return result;
//	}

//	void Range::CreateNames(const VARIANT& Top, const VARIANT& Left, const VARIANT& Bottom, const VARIANT& Right)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT;
//		InvokeHelper(0x1c9, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 &Top, &Left, &Bottom, &Right);
//	}
//
//	void Range::CreatePublisher(const VARIANT& Edition, long Appearance, const VARIANT& ContainsPICT, const VARIANT& ContainsBIFF, const VARIANT& ContainsRTF, const VARIANT& ContainsVALU)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT VTS_I4 VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT;
//		InvokeHelper(0x1ca, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 &Edition, Appearance, &ContainsPICT, &ContainsBIFF, &ContainsRTF, &ContainsVALU);
//	}
//
//	LPDISPATCH Range::GetCurrentArray()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x1f5, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}
//
//	LPDISPATCH Range::GetCurrentRegion()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0xf3, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}
//
//	void Range::Cut(const VARIANT& Destination)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0x235, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 &Destination);
//	}
//
//	void Range::DataSeries(const VARIANT& Rowcol, long Type, long Date, const VARIANT& Step, const VARIANT& Stop, const VARIANT& Trend)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT VTS_I4 VTS_I4 VTS_VARIANT VTS_VARIANT VTS_VARIANT;
//		InvokeHelper(0x1d0, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 &Rowcol, Type, Date, &Step, &Stop, &Trend);
//	}
//
//	VARIANT Range::Get_Default(const VARIANT& RowIndex, const VARIANT& ColumnIndex)
//	{
//		VARIANT result;
//		static BYTE parms[] =
//			VTS_VARIANT VTS_VARIANT;
//		InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, parms,
//			&RowIndex, &ColumnIndex);
//		return result;
//	}
//
//	void Range::Set_Default(const VARIANT& RowIndex, const VARIANT& ColumnIndex, const VARIANT& newValue)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT VTS_VARIANT VTS_VARIANT;
//		InvokeHelper(0x0, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 &RowIndex, &ColumnIndex, &newValue);
//	}
//
//	void Range::Delete(const VARIANT& Shift)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0x75, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 &Shift);
//	}
//
//	LPDISPATCH Range::GetDependents()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x21f, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}
//
//	VARIANT Range::DialogBox_()
//	{
//		VARIANT result;
//		InvokeHelper(0xf5, DISPATCH_METHOD, VT_VARIANT, (void*)&result, NULL);
//		return result;
//	}
//
//	LPDISPATCH Range::GetDirectDependents()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x221, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}
//
//	LPDISPATCH Range::GetDirectPrecedents()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x222, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}
//
//	VARIANT Range::EditionOptions(long Type, long Option, const VARIANT& Name, const VARIANT& Reference, long Appearance, long ChartSize, const VARIANT& Format)
//	{
//		VARIANT result;
//		static BYTE parms[] =
//			VTS_I4 VTS_I4 VTS_VARIANT VTS_VARIANT VTS_I4 VTS_I4 VTS_VARIANT;
//		InvokeHelper(0x46b, DISPATCH_METHOD, VT_VARIANT, (void*)&result, parms,
//			Type, Option, &Name, &Reference, Appearance, ChartSize, &Format);
//		return result;
//	}
//
//	LPDISPATCH Range::GetEnd(long Direction)
//	{
//		LPDISPATCH result;
//		static BYTE parms[] =
//			VTS_I4;
//		InvokeHelper(0x1f4, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, parms,
//			Direction);
//		return result;
//	}
//
//	LPDISPATCH Range::GetEntireColumn()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0xf6, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}
//
//	LPDISPATCH Range::GetEntireRow()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0xf7, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}
//
//	void Range::FillDown()
//	{
//		InvokeHelper(0xf8, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
//	}
//
//	void Range::FillLeft()
//	{
//		InvokeHelper(0xf9, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
//	}
//
//	void Range::FillRight()
//	{
//		InvokeHelper(0xfa, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
//	}
//
//	void Range::FillUp()
//	{
//		InvokeHelper(0xfb, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
//	}
//
//	LPDISPATCH Range::Find(const VARIANT& What, const VARIANT& After, const VARIANT& LookIn, const VARIANT& LookAt, const VARIANT& SearchOrder, long SearchDirection, const VARIANT& MatchCase, const VARIANT& MatchByte, 
//			const VARIANT& MatchControlCharacters, const VARIANT& MatchDiacritics, const VARIANT& MatchKashida, const VARIANT& MatchAlefHamza)
//	{
//		LPDISPATCH result;
//		static BYTE parms[] =
//			VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_I4 VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT;
//		InvokeHelper(0x18e, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms,
//			&What, &After, &LookIn, &LookAt, &SearchOrder, SearchDirection, &MatchCase, &MatchByte, &MatchControlCharacters, &MatchDiacritics, &MatchKashida, &MatchAlefHamza);
//		return result;
//	}
//
//	LPDISPATCH Range::FindNext(const VARIANT& After)
//	{
//		LPDISPATCH result;
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0x18f, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms,
//			&After);
//		return result;
//	}
//
//	LPDISPATCH Range::FindPrevious(const VARIANT& After)
//	{
//		LPDISPATCH result;
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0x190, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms,
//			&After);
//		return result;
//	}

	// 2001-03-19 confirmed.	<JP>動作確認済み</JP>
	public ExcelFont Font() throws JComException { return new ExcelFont((IDispatch)get("Font")); }
//	LPDISPATCH Range::GetFont()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x92, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}

	public String Formula() throws JComException { return (String)get("Formula"); }
//	VARIANT Range::GetFormula()
//	{
//		VARIANT result;
//		InvokeHelper(0x105, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
//		return result;
//	}

	// 式を"=SUM(B1:B56)"の形式で指定する。動作確認済み		
	public void Formula(String newValue) throws JComException { put("Formula", newValue); }
//	void Range::SetFormula(const VARIANT& newValue)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0x105, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 &newValue);
//	}
//
//	VARIANT Range::GetFormulaArray()
//	{
//		VARIANT result;
//		InvokeHelper(0x24a, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
//		return result;
//	}
//
//	void Range::SetFormulaArray(const VARIANT& newValue)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0x24a, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 &newValue);
//	}
//
//	long Range::GetFormulaLabel()
//	{
//		long result;
//		InvokeHelper(0x564, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
//		return result;
//	}
//
//	void Range::SetFormulaLabel(long nNewValue)
//	{
//		static BYTE parms[] =
//			VTS_I4;
//		InvokeHelper(0x564, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 nNewValue);
//	}
//
//	VARIANT Range::GetFormulaHidden()
//	{
//		VARIANT result;
//		InvokeHelper(0x106, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
//		return result;
//	}
//
//	void Range::SetFormulaHidden(const VARIANT& newValue)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0x106, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 &newValue);
//	}
//
//	VARIANT Range::GetFormulaLocal()
//	{
//		VARIANT result;
//		InvokeHelper(0x107, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
//		return result;
//	}
//
//	void Range::SetFormulaLocal(const VARIANT& newValue)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0x107, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 &newValue);
//	}
//
//	VARIANT Range::GetFormulaR1C1()
//	{
//		VARIANT result;
//		InvokeHelper(0x108, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
//		return result;
//	}
//
//	void Range::SetFormulaR1C1(const VARIANT& newValue)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0x108, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 &newValue);
//	}
//
//	VARIANT Range::GetFormulaR1C1Local()
//	{
//		VARIANT result;
//		InvokeHelper(0x109, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
//		return result;
//	}
//
//	void Range::SetFormulaR1C1Local(const VARIANT& newValue)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0x109, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 &newValue);
//	}
//
//	void Range::FunctionWizard()
//	{
//		InvokeHelper(0x23b, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
//	}
//
//	BOOL Range::GoalSeek(const VARIANT& Goal, LPDISPATCH ChangingCell)
//	{
//		BOOL result;
//		static BYTE parms[] =
//			VTS_VARIANT VTS_DISPATCH;
//		InvokeHelper(0x1d8, DISPATCH_METHOD, VT_BOOL, (void*)&result, parms,
//			&Goal, ChangingCell);
//		return result;
//	}
//
//	VARIANT Range::Group(const VARIANT& Start, const VARIANT& End, const VARIANT& By, const VARIANT& Periods)
//	{
//		VARIANT result;
//		static BYTE parms[] =
//			VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT;
//		InvokeHelper(0x2e, DISPATCH_METHOD, VT_VARIANT, (void*)&result, parms,
//			&Start, &End, &By, &Periods);
//		return result;
//	}
//
//	VARIANT Range::GetHasArray()
//	{
//		VARIANT result;
//		InvokeHelper(0x10a, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
//		return result;
//	}
//
//	VARIANT Range::GetHasFormula()
//	{
//		VARIANT result;
//		InvokeHelper(0x10b, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
//		return result;
//	}
//
//	VARIANT Range::GetHeight()
//	{
//		VARIANT result;
//		InvokeHelper(0x7b, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
//		return result;
//	}
//
//	VARIANT Range::GetHidden()
//	{
//		VARIANT result;
//		InvokeHelper(0x10c, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
//		return result;
//	}
//
//	void Range::SetHidden(const VARIANT& newValue)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0x10c, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 &newValue);
//	}
//
//	VARIANT Range::GetHorizontalAlignment()
//	{
//		VARIANT result;
//		InvokeHelper(0x88, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
//		return result;
//	}
//
//	void Range::SetHorizontalAlignment(const VARIANT& newValue)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0x88, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 &newValue);
//	}
//
//	VARIANT Range::GetIndentLevel()
//	{
//		VARIANT result;
//		InvokeHelper(0xc9, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
//		return result;
//	}
//
//	void Range::SetIndentLevel(const VARIANT& newValue)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0xc9, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 &newValue);
//	}
//
//	void Range::InsertIndent(long InsertAmount)
//	{
//		static BYTE parms[] =
//			VTS_I4;
//		InvokeHelper(0x565, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 InsertAmount);
//	}
//
//	void Range::Insert(const VARIANT& Shift)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0xfc, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 &Shift);
//	}
//
//	LPDISPATCH Range::GetInterior()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x81, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}


	// 指定したセルを取得する。 1オリジンなので注意
	public ExcelRange Item(int RowIndex, int ColumnIndex) throws JComException {
		Object[] arglist = new Object[2];
		arglist[0] = new Integer(RowIndex);
		arglist[1] = new Integer(ColumnIndex);
		return new ExcelRange((IDispatch)get("Item", arglist));
	}
	public ExcelRange Item(int RowIndex) throws JComException {
		Object[] arglist = new Object[1];
		arglist[0] = new Integer(RowIndex);
		return new ExcelRange((IDispatch)get("Item", arglist));
	}
//	VARIANT Range::GetItem(const VARIANT& RowIndex, const VARIANT& ColumnIndex)
//	{
//		VARIANT result;
//		static BYTE parms[] =
//			VTS_VARIANT VTS_VARIANT;
//		InvokeHelper(0xaa, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, parms,
//			&RowIndex, &ColumnIndex);
//		return result;
//	}
//
//	void Range::SetItem(const VARIANT& RowIndex, const VARIANT& ColumnIndex, const VARIANT& newValue)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT VTS_VARIANT VTS_VARIANT;
//		InvokeHelper(0xaa, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 &RowIndex, &ColumnIndex, &newValue);
//	}
//
//	void Range::Justify()
//	{
//		InvokeHelper(0x1ef, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
//	}
//
//	VARIANT Range::GetLeft()
//	{
//		VARIANT result;
//		InvokeHelper(0x7f, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
//		return result;
//	}
//
//	long Range::GetListHeaderRows()
//	{
//		long result;
//		InvokeHelper(0x4a3, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
//		return result;
//	}
//
//	void Range::ListNames()
//	{
//		InvokeHelper(0xfd, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
//	}
//
//	long Range::GetLocationInTable()
//	{
//		long result;
//		InvokeHelper(0x2b3, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
//		return result;
//	}
//
//	VARIANT Range::GetLocked()
//	{
//		VARIANT result;
//		InvokeHelper(0x10d, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
//		return result;
//	}
//
//	void Range::SetLocked(const VARIANT& newValue)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0x10d, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 &newValue);
//	}
//
//	void Range::Merge(const VARIANT& Across)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0x234, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 &Across);
//	}
//
//	void Range::UnMerge()
//	{
//		InvokeHelper(0x568, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
//	}
//
//	LPDISPATCH Range::GetMergeArea()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x569, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}
//
//	VARIANT Range::GetMergeCells()
//	{
//		VARIANT result;
//		InvokeHelper(0xd0, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
//		return result;
//	}
//
//	void Range::SetMergeCells(const VARIANT& newValue)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0xd0, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 &newValue);
//	}
//
//	VARIANT Range::GetName()
//	{
//		VARIANT result;
//		InvokeHelper(0x6e, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
//		return result;
//	}
//
//	void Range::SetName(const VARIANT& newValue)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0x6e, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 &newValue);
//	}
//
//	void Range::NavigateArrow(const VARIANT& TowardPrecedent, const VARIANT& ArrowNumber, const VARIANT& LinkNumber)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT VTS_VARIANT VTS_VARIANT;
//		InvokeHelper(0x408, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 &TowardPrecedent, &ArrowNumber, &LinkNumber);
//	}
//
//	LPUNKNOWN Range::Get_NewEnum()
//	{
//		LPUNKNOWN result;
//		InvokeHelper(0xfffffffc, DISPATCH_PROPERTYGET, VT_UNKNOWN, (void*)&result, NULL);
//		return result;
//	}
//
//	LPDISPATCH Range::GetNext()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x1f6, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}
//
//	CString Range::NoteText(const VARIANT& Text, const VARIANT& Start, const VARIANT& Length)
//	{
//		CString result;
//		static BYTE parms[] =
//			VTS_VARIANT VTS_VARIANT VTS_VARIANT;
//		InvokeHelper(0x467, DISPATCH_METHOD, VT_BSTR, (void*)&result, parms,
//			&Text, &Start, &Length);
//		return result;
//	}
//
//	VARIANT Range::GetNumberFormat()
//	{
//		VARIANT result;
//		InvokeHelper(0xc1, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
//		return result;
//	}
//
//	void Range::SetNumberFormat(const VARIANT& newValue)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0xc1, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 &newValue);
//	}
//
//	VARIANT Range::GetNumberFormatLocal()
//	{
//		VARIANT result;
//		InvokeHelper(0x449, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
//		return result;
//	}
//
//	void Range::SetNumberFormatLocal(const VARIANT& newValue)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0x449, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 &newValue);
//	}
//
//	LPDISPATCH Range::GetOffset(const VARIANT& RowOffset, const VARIANT& ColumnOffset)
//	{
//		LPDISPATCH result;
//		static BYTE parms[] =
//			VTS_VARIANT VTS_VARIANT;
//		InvokeHelper(0xfe, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, parms,
//			&RowOffset, &ColumnOffset);
//		return result;
//	}
//
//	VARIANT Range::GetOrientation()
//	{
//		VARIANT result;
//		InvokeHelper(0x86, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
//		return result;
//	}
//
//	void Range::SetOrientation(const VARIANT& newValue)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0x86, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 &newValue);
//	}
//
//	VARIANT Range::GetOutlineLevel()
//	{
//		VARIANT result;
//		InvokeHelper(0x10f, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
//		return result;
//	}
//
//	void Range::SetOutlineLevel(const VARIANT& newValue)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0x10f, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 &newValue);
//	}


//	long Range::GetPageBreak()
//	{
//		long result;
//		InvokeHelper(0xff, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
//		return result;
//	}

	// 改ページ位置を指定します。XlPageBreakオブジェクトの定数を指定してください。
	public void PageBreak(int newValue) throws JComException { put("PageBreak", new Integer(newValue)); }
//	void Range::SetPageBreak(long nNewValue)
//	{
//		static BYTE parms[] =
//			VTS_I4;
//		InvokeHelper(0xff, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 nNewValue);
//	}
//
//	void Range::Parse(const VARIANT& ParseLine, const VARIANT& Destination)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT VTS_VARIANT;
//		InvokeHelper(0x1dd, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 &ParseLine, &Destination);
//	}
//
//	void Range::PasteSpecial(long Paste, long Operation, const VARIANT& SkipBlanks, const VARIANT& Transpose)
//	{
//		static BYTE parms[] =
//			VTS_I4 VTS_I4 VTS_VARIANT VTS_VARIANT;
//		InvokeHelper(0x403, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 Paste, Operation, &SkipBlanks, &Transpose);
//	}
//
//	LPDISPATCH Range::GetPivotField()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x2db, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}
//
//	LPDISPATCH Range::GetPivotItem()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x2e4, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}
//
//	LPDISPATCH Range::GetPivotTable()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x2cc, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}
//
//	LPDISPATCH Range::GetPrecedents()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x220, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}
//
//	VARIANT Range::GetPrefixCharacter()
//	{
//		VARIANT result;
//		InvokeHelper(0x1f8, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
//		return result;
//	}
//
//	LPDISPATCH Range::GetPrevious()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x1f7, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}

	// 動作確認済み
	public void PrintOut() throws JComException {
		method("PrintOut", null);
	}
//	void Range::PrintOut(const VARIANT& From, const VARIANT& To, const VARIANT& Copies, const VARIANT& Preview, const VARIANT& ActivePrinter, const VARIANT& PrintToFile, const VARIANT& Collate)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT;
//		InvokeHelper(0x389, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 &From, &To, &Copies, &Preview, &ActivePrinter, &PrintToFile, &Collate);
//	}
//
//	void Range::PrintPreview(const VARIANT& EnableChanges)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0x119, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 &EnableChanges);
//	}
//
//	LPDISPATCH Range::GetQueryTable()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x56a, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}
//
//	LPDISPATCH Range::GetRange(const VARIANT& Cell1, const VARIANT& Cell2)
//	{
//		LPDISPATCH result;
//		static BYTE parms[] =
//			VTS_VARIANT VTS_VARIANT;
//		InvokeHelper(0xc5, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, parms,
//			&Cell1, &Cell2);
//		return result;
//	}
//
//	void Range::RemoveSubtotal()
//	{
//		InvokeHelper(0x373, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
//	}
//
//	BOOL Range::Replace(const VARIANT& What, const VARIANT& Replacement, const VARIANT& LookAt, const VARIANT& SearchOrder, const VARIANT& MatchCase, const VARIANT& MatchByte, const VARIANT& MatchControlCharacters, const VARIANT& MatchDiacritics, 
//			const VARIANT& MatchKashida, const VARIANT& MatchAlefHamza)
//	{
//		BOOL result;
//		static BYTE parms[] =
//			VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT;
//		InvokeHelper(0xe2, DISPATCH_METHOD, VT_BOOL, (void*)&result, parms,
//			&What, &Replacement, &LookAt, &SearchOrder, &MatchCase, &MatchByte, &MatchControlCharacters, &MatchDiacritics, &MatchKashida, &MatchAlefHamza);
//		return result;
//	}
//
//	LPDISPATCH Range::GetResize(const VARIANT& RowSize, const VARIANT& ColumnSize)
//	{
//		LPDISPATCH result;
//		static BYTE parms[] =
//			VTS_VARIANT VTS_VARIANT;
//		InvokeHelper(0x100, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, parms,
//			&RowSize, &ColumnSize);
//		return result;
//	}
//
//	long Range::GetRow()
//	{
//		long result;
//		InvokeHelper(0x101, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
//		return result;
//	}
//
//	LPDISPATCH Range::RowDifferences(const VARIANT& Comparison)
//	{
//		LPDISPATCH result;
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0x1ff, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms,
//			&Comparison);
//		return result;
//	}
//
//	VARIANT Range::GetRowHeight()
//	{
//		VARIANT result;
//		InvokeHelper(0x110, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
//		return result;
//	}
//
//	void Range::SetRowHeight(const VARIANT& newValue)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0x110, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 &newValue);
//	}
//

	public ExcelRange Rows() throws JComException { return new ExcelRange((IDispatch)get("Rows")); }
//	LPDISPATCH Range::GetRows()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x102, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}
//
//	VARIANT Range::Run(const VARIANT& Arg1, const VARIANT& Arg2, const VARIANT& Arg3, const VARIANT& Arg4, const VARIANT& Arg5, const VARIANT& Arg6, const VARIANT& Arg7, const VARIANT& Arg8, const VARIANT& Arg9, const VARIANT& Arg10, 
//			const VARIANT& Arg11, const VARIANT& Arg12, const VARIANT& Arg13, const VARIANT& Arg14, const VARIANT& Arg15, const VARIANT& Arg16, const VARIANT& Arg17, const VARIANT& Arg18, const VARIANT& Arg19, const VARIANT& Arg20, 
//			const VARIANT& Arg21, const VARIANT& Arg22, const VARIANT& Arg23, const VARIANT& Arg24, const VARIANT& Arg25, const VARIANT& Arg26, const VARIANT& Arg27, const VARIANT& Arg28, const VARIANT& Arg29, const VARIANT& Arg30)
//	{
//		VARIANT result;
//		static BYTE parms[] =
//			VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT 
//			VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT;
//		InvokeHelper(0x103, DISPATCH_METHOD, VT_VARIANT, (void*)&result, parms,
//			&Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30);
//		return result;
//	}
//
//	void Range::Select()
//	{
//		InvokeHelper(0xeb, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
//	}
//
//	void Range::Show()
//	{
//		InvokeHelper(0x1f0, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
//	}
//
//	void Range::ShowDependents(const VARIANT& Remove)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0x36d, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 &Remove);
//	}
//
//	VARIANT Range::GetShowDetail()
//	{
//		VARIANT result;
//		InvokeHelper(0x249, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
//		return result;
//	}
//
//	void Range::SetShowDetail(const VARIANT& newValue)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0x249, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 &newValue);
//	}
//
//	void Range::ShowErrors()
//	{
//		InvokeHelper(0x36e, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
//	}
//
//	void Range::ShowPrecedents(const VARIANT& Remove)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0x36f, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 &Remove);
//	}
//
//	VARIANT Range::GetShrinkToFit()
//	{
//		VARIANT result;
//		InvokeHelper(0xd1, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
//		return result;
//	}
//
//	void Range::SetShrinkToFit(const VARIANT& newValue)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0xd1, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 &newValue);
//	}
//
//	void Range::Sort(const VARIANT& Key1, long Order1, const VARIANT& Key2, const VARIANT& Type, long Order2, const VARIANT& Key3, long Order3, long Header, const VARIANT& OrderCustom, const VARIANT& MatchCase, long Orientation, long SortMethod, 
//			const VARIANT& IgnoreControlCharacters, const VARIANT& IgnoreDiacritics, const VARIANT& IgnoreKashida)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT VTS_I4 VTS_VARIANT VTS_VARIANT VTS_I4 VTS_VARIANT VTS_I4 VTS_I4 VTS_VARIANT VTS_VARIANT VTS_I4 VTS_I4 VTS_VARIANT VTS_VARIANT VTS_VARIANT;
//		InvokeHelper(0x370, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 &Key1, Order1, &Key2, &Type, Order2, &Key3, Order3, Header, &OrderCustom, &MatchCase, Orientation, SortMethod, &IgnoreControlCharacters, &IgnoreDiacritics, &IgnoreKashida);
//	}
//
//	void Range::SortSpecial(long SortMethod, const VARIANT& Key1, long Order1, const VARIANT& Type, const VARIANT& Key2, long Order2, const VARIANT& Key3, long Order3, long Header, const VARIANT& OrderCustom, const VARIANT& MatchCase, long Orientation)//
	{//
	//	static BYTE parms[] =
	//		VTS_I4 VTS_VARIANT VTS_I4 VTS_VARIANT VTS_VARIANT VTS_I4 VTS_VARIANT VTS_I4 VTS_I4 VTS_VARIANT VTS_VARIANT VTS_I4;
	//	InvokeHelper(0x371, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
	//		 SortMethod, &Key1, Order1, &Type, &Key2, Order2, &Key3, Order3, Header, &OrderCustom, &MatchCase, Orientation);
	}//
//
//	LPDISPATCH Range::GetSoundNote()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x394, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}
//
//	LPDISPATCH Range::SpecialCells(long Type, const VARIANT& Value)
//	{
//		LPDISPATCH result;
//		static BYTE parms[] =
//			VTS_I4 VTS_VARIANT;
//		InvokeHelper(0x19a, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms,
//			Type, &Value);
//		return result;
//	}
//
//	VARIANT Range::GetStyle()
//	{
//		VARIANT result;
//		InvokeHelper(0x104, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
//		return result;
//	}
//
//	void Range::SetStyle(const VARIANT& newValue)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0x104, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 &newValue);
//	}
//
//	void Range::SubscribeTo(LPCTSTR Edition, long Format)
//	{
//		static BYTE parms[] =
//			VTS_BSTR VTS_I4;
//		InvokeHelper(0x1e1, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 Edition, Format);
//	}
//
//	void Range::Subtotal(long GroupBy, long Function, const VARIANT& TotalList, const VARIANT& Replace, const VARIANT& PageBreaks, long SummaryBelowData)
//	{
//		static BYTE parms[] =
//			VTS_I4 VTS_I4 VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_I4;
//		InvokeHelper(0x372, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 GroupBy, Function, &TotalList, &Replace, &PageBreaks, SummaryBelowData);
//	}
//
//	VARIANT Range::GetSummary()
//	{
//		VARIANT result;
//		InvokeHelper(0x111, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
//		return result;
//	}
//
//	void Range::Table(const VARIANT& RowInput, const VARIANT& ColumnInput)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT VTS_VARIANT;
//		InvokeHelper(0x1f1, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 &RowInput, &ColumnInput);
//	}

	public String Text() throws JComException { return (String)get("Text"); }
//	VARIANT Range::GetText()
//	{
//		VARIANT result;
//		InvokeHelper(0x8a, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
//		return result;
//	}

//	void Range::TextToColumns(const VARIANT& Destination, long DataType, long TextQualifier, const VARIANT& ConsecutiveDelimiter, const VARIANT& Tab, const VARIANT& Semicolon, const VARIANT& Comma, const VARIANT& Space, const VARIANT& Other, 
//			const VARIANT& OtherChar, const VARIANT& FieldInfo)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT VTS_I4 VTS_I4 VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT;
//		InvokeHelper(0x410, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
//			 &Destination, DataType, TextQualifier, &ConsecutiveDelimiter, &Tab, &Semicolon, &Comma, &Space, &Other, &OtherChar, &FieldInfo);
//	}
//
//	VARIANT Range::GetTop()
//	{
//		VARIANT result;
//		InvokeHelper(0x7e, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
//		return result;
//	}
//
//	void Range::Ungroup()
//	{
//		InvokeHelper(0xf4, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
//	}
//
//	VARIANT Range::GetUseStandardHeight()
//	{
//		VARIANT result;
//		InvokeHelper(0x112, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
//		return result;
//	}
//
//	void Range::SetUseStandardHeight(const VARIANT& newValue)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0x112, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 &newValue);
//	}
//
//	VARIANT Range::GetUseStandardWidth()
//	{
//		VARIANT result;
//		InvokeHelper(0x113, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
//		return result;
//	}
//
//	void Range::SetUseStandardWidth(const VARIANT& newValue)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0x113, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 &newValue);
//	}
//
//	LPDISPATCH Range::GetValidation()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x56b, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}

	// セルの値を取得します。戻り値は以下の表を参考にしてください。
	public Object Value() throws JComException { return get("Value"); }
//	VARIANT Range::GetValue()
//	{
//		VARIANT result;
//		InvokeHelper(0x6, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
//		return result;
//	}

	// セルに値をセットします。すべて動作確認済み
	//	セルの書式設定とVARIANT型とJava型の対応は以下のとおりです。
	//	空				VT_EMPTY	null
	//	標準			VT_BSTR		String
	//	数値			VT_R8		double
	//	通貨			VT_CY		VariantCurrency
	//	会計			VT_CY		VariantCurrency
	//	日時			VT_DATE		Date
	//	時刻			VT_DATE		Date
	//	パーセンテージ	VT_R8		double
	//	分数			VT_R8		double
	//	指数			VT_R8		double
	//	文字列			VT_BSTR		String
	//	その他			
	//	ユーザー定義	
	public void Value(double newValue) throws JComException { put("Value", new Double(newValue)); }
	public void Value(String newValue) throws JComException { put("Value", newValue); }
	public void Value(Date newValue) throws JComException { put("Value", newValue); }
	public void Value(VariantCurrency newValue) throws JComException { put("Value", newValue); }
//	void Range::SetValue(const VARIANT& newValue)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0x6, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 &newValue);
//	}
//
//	VARIANT Range::GetValue2()
//	{
//		VARIANT result;
//		InvokeHelper(0x56c, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
//		return result;
//	}
//
//	void Range::SetValue2(const VARIANT& newValue)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0x56c, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 &newValue);
//	}
//
//	VARIANT Range::GetVerticalAlignment()
//	{
//		VARIANT result;
//		InvokeHelper(0x89, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
//		return result;
//	}
//
//	void Range::SetVerticalAlignment(const VARIANT& newValue)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0x89, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 &newValue);
//	}
//
//	VARIANT Range::GetWidth()
//	{
//		VARIANT result;
//		InvokeHelper(0x7a, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
//		return result;
//	}
//

	public ExcelWorksheet Worksheet() throws JComException { return new ExcelWorksheet((IDispatch)get("Worksheet")); }
//	LPDISPATCH Range::GetWorksheet()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x15c, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}

//	VARIANT Range::GetWrapText()
//	{
//		VARIANT result;
//		InvokeHelper(0x114, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
//		return result;
//	}
//
//	void Range::SetWrapText(const VARIANT& newValue)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0x114, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 &newValue);
//	}
//
//	LPDISPATCH Range::AddComment(const VARIANT& Text)
//	{
//		LPDISPATCH result;
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0x56d, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms,
//			&Text);
//		return result;
//	}
//
//	LPDISPATCH Range::GetComment()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x38e, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}
//
//	void Range::ClearComments()
//	{
//		InvokeHelper(0x56e, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
//	}
//
//	LPDISPATCH Range::GetPhonetic()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x56f, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}
//
//	LPDISPATCH Range::GetFormatConditions()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x570, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}
//
//	long Range::GetReadingOrder()
//	{
//		long result;
//		InvokeHelper(0x3cf, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
//		return result;
//	}
//
//	void Range::SetReadingOrder(long nNewValue)
//	{
//		static BYTE parms[] =
//			VTS_I4;
//		InvokeHelper(0x3cf, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 nNewValue);
//	}
//
//	LPDISPATCH Range::GetHyperlinks()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x571, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}

	// セルのとりうる型のテスト
	public static void main(String[] args) throws JComException, java.io.IOException {
		ReleaseManager rm = new ReleaseManager();
		try {
			System.out.println("Excelを起動します");
			ExcelApplication excel = new ExcelApplication(rm);
			excel.Visible(true);

			excel.Workbooks().Add();	// 新しいブックを作成
			excel.ActiveSheet().Range("B1").Value(new VariantCurrency(56000));
			excel.ActiveSheet().Range("B2").Value(new Date());	// 現在時刻
			excel.ActiveSheet().Range("B3").Value("文字列");
			excel.ActiveSheet().Range("B4").Value(12.3);
			excel.ActiveSheet().Range("B6").Value("");
			excel.ActiveSheet().Cells().Columns().AutoFit();	// 横幅をフィットさせる

			System.out.println("A1のセルに値を設定してください。");
			System.out.println("設定したら[Enter]を押してください");
			System.in.read();

			ExcelRange xlRange = excel.ActiveSheet().Range("A1");
			System.out.println("A1の値="+xlRange.Value());
			excel.ActiveWorkbook().Close(false,null,false);	// 保存せずに閉じる
			excel.Quit();
		} finally { rm.release(); }
	}

}
