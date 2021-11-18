package jp.ne.so_net.ga2.no_ji.jcom.excel8;
import java.lang.*;
import java.lang.reflect.*;
import jp.ne.so_net.ga2.no_ji.jcom.*;

public class ExcelFont extends IDispatch {

	public ExcelFont(IDispatch jcom) { super(jcom); }

	public ExcelApplication Application() throws JComException { return new ExcelApplication((IDispatch)get("Application")); }
//	LPDISPATCH Font::GetApplication()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x94, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}

	public int Creator() throws JComException { return ((Integer)get("Creator")).intValue(); }
//	long Font::GetCreator()
//	{
//		long result;
//		InvokeHelper(0x95, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
//		return result;
//	}

//	LPDISPATCH Font::GetParent()
//	{
//		LPDISPATCH result;
//		InvokeHelper(0x96, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
//		return result;
//	}

	public int Background() throws JComException { return ((Integer)get("Background")).intValue(); }
//	VARIANT Font::GetBackground()
//	{
//		VARIANT result;
//		InvokeHelper(0xb4, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
//		return result;
//	}

	public void BackGround(int newValue) throws JComException { put("Background", new Integer(newValue)); }
//	void Font::SetBackground(const VARIANT& newValue)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0xb4, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 &newValue);
//	}

	public boolean Bold() throws JComException { return ((Boolean)get("Bold")).booleanValue(); }
//	VARIANT Font::GetBold()
//	{
//		VARIANT result;
//		InvokeHelper(0x60, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
//		return result;
//	}

	public void Bold(boolean newValue) throws JComException { put("Bold", new Boolean(newValue)); }
//	void Font::SetBold(const VARIANT& newValue)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0x60, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 &newValue);
//	}

	public int Color() throws JComException { return ((Integer)get("Color")).intValue(); }
//	VARIANT Font::GetColor()
//	{
//		VARIANT result;
//		InvokeHelper(0x63, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
//		return result;
//	}

	public void Color(int newValue) throws JComException { put("Color", new Integer(newValue)); }
//	void Font::SetColor(const VARIANT& newValue)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0x63, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 &newValue);
//	}

	public int ColorIndex() throws JComException { return ((Integer)get("ColorIndex")).intValue(); }
//	VARIANT Font::GetColorIndex()
//	{
//		VARIANT result;
//		InvokeHelper(0x61, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
//		return result;
//	}

	public void ColorIndex(int newValue) throws JComException { put("ColorIndex", new Integer(newValue)); }
//	void Font::SetColorIndex(const VARIANT& newValue)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0x61, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 &newValue);
//	}

	public String FontStyle() throws JComException { return (String)get("FontStyle"); }
//	VARIANT Font::GetFontStyle()
//	{
//		VARIANT result;
//		InvokeHelper(0xb1, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
//		return result;
//	}

	public void FontStyle(String newValue) throws JComException { put("FontStyle", newValue); }
//	void Font::SetFontStyle(const VARIANT& newValue)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0xb1, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 &newValue);
//	}

	public boolean Italic() throws JComException { return ((Boolean)get("Italic")).booleanValue(); }
//	VARIANT Font::GetItalic()
//	{
//		VARIANT result;
//		InvokeHelper(0x65, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
//		return result;
//	}
//	
	public void Italic(boolean newValue) throws JComException { put("Italic", new Boolean(newValue)); }
//	void Font::SetItalic(const VARIANT& newValue)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0x65, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 &newValue);
//	}
//	
	public String Name() throws JComException { return (String)get("Name"); }
//	VARIANT Font::GetName()
//	{
//		VARIANT result;
//		InvokeHelper(0x6e, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
//		return result;
//	}

	public void Name(String newValue) throws JComException { put("Name", newValue); }
//	void Font::SetName(const VARIANT& newValue)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0x6e, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 &newValue);
//	}

	public boolean OutlineFont() throws JComException { return ((Boolean)get("OutlineFont")).booleanValue(); }
//	VARIANT Font::GetOutlineFont()
//	{
//		VARIANT result;
//		InvokeHelper(0xdd, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
//		return result;
//	}
//	
	public void OutlineFont(boolean newValue) throws JComException { put("OutlineFont", new Boolean(newValue)); }
//	void Font::SetOutlineFont(const VARIANT& newValue)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0xdd, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 &newValue);
//	}
//	
	public boolean Shadow() throws JComException { return ((Boolean)get("Shadow")).booleanValue(); }
//	VARIANT Font::GetShadow()
//	{
//		VARIANT result;
//		InvokeHelper(0x67, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
//		return result;
//	}
//	
	public void Shadow(boolean newValue) throws JComException { put("Shadow", new Boolean(newValue)); }
//	void Font::SetShadow(const VARIANT& newValue)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0x67, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 &newValue);
//	}


	public double Size() throws JComException { return ((Double)get("Size")).doubleValue(); }
//	VARIANT Font::GetSize()
//	{
//		VARIANT result;
//		InvokeHelper(0x68, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
//		return result;
//	}

	public void Size(double newValue) throws JComException { put("Size", new Double(newValue)); }
//	void Font::SetSize(const VARIANT& newValue)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0x68, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 &newValue);
//	}
//	
	public boolean Strikethrough() throws JComException { return ((Boolean)get("Strikethrough")).booleanValue(); }
//	VARIANT Font::GetStrikethrough()
//	{
//		VARIANT result;
//		InvokeHelper(0x69, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
//		return result;
//	}
//	
	public void Strikethrough(boolean newValue) throws JComException { put("Strikethrough", new Boolean(newValue)); }
//	void Font::SetStrikethrough(const VARIANT& newValue)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0x69, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 &newValue);
//	}

	//	‰º•t‚«•¶Žš
	public boolean Subscript() throws JComException { return ((Boolean)get("Subscript")).booleanValue(); }
//	VARIANT Font::GetSubscript()
//	{
//		VARIANT result;
//		InvokeHelper(0xb3, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
//		return result;
//	}
//	
	public void Subscript(boolean newValue) throws JComException { put("Subscript", new Boolean(newValue)); }
//	void Font::SetSubscript(const VARIANT& newValue)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0xb3, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 &newValue);
//	}
//	
	public boolean Superscript() throws JComException { return ((Boolean)get("Superscript")).booleanValue(); }
//	VARIANT Font::GetSuperscript()
//	{
//		VARIANT result;
//		InvokeHelper(0xb2, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
//		return result;
//	}
//	
	public void Superscript(boolean newValue) throws JComException { put("Superscript", new Boolean(newValue)); }
//	void Font::SetSuperscript(const VARIANT& newValue)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0xb2, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 &newValue);
//	}
//	
	public boolean Underline() throws JComException { return ((Boolean)get("Underline")).booleanValue(); }
//	VARIANT Font::GetUnderline()
//	{
//		VARIANT result;
//		InvokeHelper(0x6a, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
//		return result;
//	}
//	
	public void Underline(boolean newValue) throws JComException { put("Underline", new Boolean(newValue)); }
//	void Font::SetUnderline(const VARIANT& newValue)
//	{
//		static BYTE parms[] =
//			VTS_VARIANT;
//		InvokeHelper(0x6a, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
//			 &newValue);
//	}

}
