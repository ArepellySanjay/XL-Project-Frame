package XL_TestCases;

import java.io.IOException;

import XL_AppUtils.XL_Utils;
import XL_AppUtils.XL_Utils_Next;

public class XL_Utils_testCase 
{

//	String xlfile ="E:\\Qedge\\XL.operation.xlsx";
	
	
	public static void main(String[] args) throws IOException
{
	
		//String xlfile ="E:\\Qedge\\XL.operation.xlsx";
	
		
		
		
	int xx =	XL_Utils_Next.getRowcount("E:\\Qedge\\XL.operation.xlsx", "EmpData");		
	System.out.println(xx);
		
	
	
	
		int ss = XL_Utils_Next.getcoloumncount("E:\\Qedge\\XL.operation.xlsx", "EmpData", 1);
		 System.out.println(ss);
		 
		 
		 
		
  String x =   XL_Utils_Next.getStringData("E:\\Qedge\\XL.operation.xlsx", "EmpData", 1, 0);
     System.out.println(x);
		
		
     
   double sa =   XL_Utils_Next.getNumaricData("E:\\Qedge\\XL.operation.xlsx", "EmpData", 2, 2);
     System.out.println(sa);
     
     
     
  boolean sanju =   XL_Utils_Next.getbooleanData("E:\\Qedge\\XL.operation.xlsx", "EmpData", 2, 3);
     System.out.println(sanju);
     
     XL_Utils_Next.getsetData("E:\\Qedge\\XL.operation.xlsx", "EmpData", 3, 4, "pass");
     
     XL_Utils_Next.getsetData("E:\\Qedge\\XL.operation.xlsx", "EmpData", 4, 4, "pass");
     
     XL_Utils_Next.getsetData("E:\\Qedge\\XL.operation.xlsx", "EmpData", 5, 4, "pass");
     
     
     
    
     XL_Utils_Next.FillGreenColour("E:\\Qedge\\XL.operation.xlsx", "EmpData", 4, 4 );     
     XL_Utils_Next.FillGreenColour("E:\\Qedge\\XL.operation.xlsx", "EmpData", 5, 4);
     
     
     XL_Utils_Next.FillRedColour("E:\\Qedge\\XL.operation.xlsx", "EmpData", 3, 4);
     
        String sanju1 = XL_Utils_Next.getStringSanjuData("E:\\Qedge\\XL.operation.xlsx", "EmpData", 6, 2);
       System.out.println(sanju1);
       
       
       
       
       
	}

}
