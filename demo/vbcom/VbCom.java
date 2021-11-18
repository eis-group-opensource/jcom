import jp.ne.so_net.ga2.no_ji.jcom.*;
import java.util.*;


/**
	It is the sample that COM made by VB is called. 
	VBで作られたＣＯＭを呼ぶサンプルです。
*/
public class VbCom {
	public static void main(String[] args) {
		ReleaseManager rm = new ReleaseManager();
	    try {
	        IDispatch vbcom = new IDispatch(rm, "Project1.Class1");

			{
			// Public Function testByte(ByVal a As Byte, ByRef b As Byte) As Byte
			Byte x = new Byte((byte)1);
			byte[] y = new byte[] { 2 };
	        Object[] param = new Object[] { x, y };
	        Byte retcode = (Byte)vbcom.method( "testByte", param );
			System.out.println("x="+x+" y="+y[0]+" retcode="+retcode);
			}

			{
			// Public Function testInteger(ByVal a As Integer, ByRef b As Integer) As Integer
			Short x = new Short((short)1);
			short[] y = new short[] { 2 };
	        Object[] param = new Object[] { x, y };
	        Short retcode = (Short)vbcom.method( "testInteger", param );
			System.out.println("x="+x+" y="+y[0]+" retcode="+retcode);
			}
			
			{
			// Public Function testLong(ByVal a As Long, ByRef b As Long) As Long
			Integer x = new Integer(1);
			int[] y = new int[] { 2 };
	        Object[] param = new Object[] { x, y };
	        Integer retcode = (Integer)vbcom.method( "testLong", param );
			System.out.println("x="+x+" y="+y[0]+" retcode="+retcode);
			}

			{
			// Public Function testSingle(ByVal a As Single, ByRef b As Single) As Single
			Float x = new Float(1.0f);
			float[] y = new float[] { 2.0f };
	        Object[] param = new Object[] { x, y };
	        Float retcode = (Float)vbcom.method( "testSingle", param );
			System.out.println("x="+x+" y="+y[0]+" retcode="+retcode);
			}

			{
			// Public Function testDouble(ByVal a As Double, ByRef b As Double) As Double
			Double x = new Double(1.0);
			double[] y = new double[] { 2.0 };
	        Object[] param = new Object[] { x, y };
	        Double retcode = (Double)vbcom.method( "testDouble", param );
			System.out.println("x="+x+" y="+y[0]+" retcode="+retcode);
			}

			{
			// Public Function testBoolean(ByVal a As Boolean, ByRef b As Boolean) As Boolean
			Boolean x = new Boolean(true);
			boolean[] y = new boolean[] { false };
	        Object[] param = new Object[] { x, y };
	        Boolean retcode = (Boolean)vbcom.method( "testBoolean", param );
			System.out.println("x="+x+" y="+y[0]+" retcode="+retcode);
			}

			{
			// Public Function testString(ByVal a As String, ByRef b As String) As String
			String x = "1";
			String[] y = new String[] { "2" };
	        Object[] param = new Object[] { x, y };
	        String retcode = (String)vbcom.method( "testString", param );
			System.out.println("x="+x+" y="+y[0]+" retcode="+retcode);
			}
/*
		JCom does not support "ByRef As Date" and "ByRef As Currency".
		You should use Long or String ... , as much as possible.
		JComは "ByRef As Date"と"ByRef As Currency"には対応していません。
		可能な限り、Long か String などを使って下さい。
			{
			// Public Function testDate(ByVal a As Date, ByRef b As Date) As Date
			GregorianCalendar calen = new GregorianCalendar();//.getInstance().getTime();
			calen.set(1967, 12, 24, 14, 27, 0);
//			Calendar y0 = new Calendar();
			Date x = calen.getTime();
			calen.set(2000, 1, 1, 0, 0, 0);
			Date[] y = new Date[] { calen.getTime() };
	        Object[] param = new Object[] { x, y };
	        Date retcode = (Date)vbcom.method( "testDate", param );
			System.out.println("x="+x+" y="+y[0]+" retcode="+retcode);
			}

			Public Function testCurrency(ByVal a As Currency, ByRef b As Currency) As Currency
			    b = a
			    testCurrency = b
			End Function
*/
			{
			// Public Sub testVoid()
			vbcom.method("testVoid", null);
			}

	    }
	    catch ( Exception e ) {
	        e.printStackTrace();
	    }
	    finally {
	        rm.release();
	    }
	}
}
/**
	An executive result
	実行結果

>java VbCom
x=1 y=1 retcode=1
x=1 y=1 retcode=1
x=1 y=1 retcode=1
x=1.0 y=1.0 retcode=1.0
x=1.0 y=1.0 retcode=1.0
x=true y=true retcode=true
x=1 y=1 retcode=1

If you fail in the call of Project1.dll, do the following command.
もし、Project1.dllの呼び出しに失敗したら、以下のコマンドを実行してください。
>regsvr32 Project1.dll

*/
