package XL_TestCases;

import java.io.IOException;

import XL_AppUtils.XL_Utils;


public class Testcase {

	public static void main(String[] args) throws IOException 	{
		
		
		String xlfile = "E:\\Qedge\\XL.operation.xlsx";
		
		

	//int x =	XL_Utils.getRowcount("E:\\Qedge\\XL.operation.xlsx", "EmpData");
	//	System.out.println(x);
		
		
	
		int x =	XL_Utils.getcoloumncount("E:\\Qedge\\XL.operation.xlsx", "LoginData", 1);
		System.out.println(x);
		
		
		
	String cellv=	XL_Utils.getStringData(xlfile,"EmpData", 2, 1);
		
		System.out.println(cellv);
		
		
		
		
    double data1 = XL_Utils.getNumericData(xlfile, "EmpData", 1, 2);
		System.out.println(data1);
		
		
		
	boolean ss =	XL_Utils.getBooleanData(xlfile, "EmpData", 1, 3);
		System.out.println(ss);
		
    
		XL_Utils.getSetData(xlfile, "EmpData", 1, 4, "Pass");
		XL_Utils.FillGreencolour(xlfile, "EmpData", 1, 4, "pass");
		
		XL_Utils.FillGreencolour(xlfile, "EmpData", 2, 4, "pass"); 
		
       XL_Utils.FillRedcolour(xlfile, "EmpData", 3, 4, "fail");
		
       XL_Utils.FillRedcolour(xlfile, "EmpData", 4, 4, "fail");
		
       XL_Utils.FillRedcolour(xlfile, "EmpData", 5, 4, "fail");
       
      // XL_Utils.FillRedcolour(xlfile, "EmpData", 6, 4, "fail");
	}

}
