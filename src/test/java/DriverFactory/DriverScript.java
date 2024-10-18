package DriverFactory;

import org.openqa.selenium.WebDriver;

import CommonFunctions.FunctionLibrary;
import utilities.ExcelFileUtile;

public class DriverScript
{
	WebDriver driver;
	String inputpath ="./FileInput/Controller.xlsx";
	String outputpath ="./FileOutput/HybridResults.xlsx";
	String TCSheet ="MasterTestCases";

	public void startTest() throws Throwable
	{
		String Module_Status="";
		String Module_New="";

		// Create Reference Project For Accessing Excel methods
		ExcelFileUtile xl = new ExcelFileUtile(inputpath);

		// Iterate All rows in TCsheet

		for(int i=1; i<=xl.rowCount(TCSheet);i++)
		{
			if(xl.getCellData(TCSheet, i, 2).equalsIgnoreCase("Y"))
			{
				//read test cases from TCSheet
				String TCModule = xl.getCellData(TCSheet, i, 1);

				// Iterate All Rows in TCModule Sheet
				for(int j=1;j<=xl.rowCount(TCModule);j++)
				{
					// Read All Cells From TcModule Sheet
					String Description = xl.getCellData(TCModule, j, 0);
					String ObjectType = xl.getCellData(TCModule, j, 1);
					String Ltype =xl.getCellData(TCModule, j, 2);
					String Lvalue = xl.getCellData(TCModule, j, 3);
					String TestData = xl.getCellData(TCModule, j, 4);

					try
					{
						if(ObjectType.equalsIgnoreCase("startBrowser"))
						{
							driver=FunctionLibrary.startBrowser();
						}

						if(ObjectType.equalsIgnoreCase("openUrl"))
						{
							FunctionLibrary.openUrl();
						}

						if(ObjectType.equalsIgnoreCase("waitForElement"))
						{
							FunctionLibrary.waitForElement(Ltype, Lvalue, TestData);
						}

						if(ObjectType.equalsIgnoreCase("typeAction"))
						{
							FunctionLibrary.typeAction(Ltype, Lvalue, TestData);
						}

						if(ObjectType.equalsIgnoreCase("clickAction"))
						{
							FunctionLibrary.clickAction(Ltype, Lvalue);
						}

						if(ObjectType.equalsIgnoreCase("validateTitle"))
						{
							FunctionLibrary.validateTitle(TestData);
						}

						if(ObjectType.equalsIgnoreCase("closeBrowser"))
						{
							FunctionLibrary.closeBrowser();
						}

						//write as pass into status cell in TCModule sheet
						xl.setCellData(TCModule, j, 5, "Pass", outputpath);
						Module_Status="True";

					}
					catch (Exception e) 
					{
						System.err.println(e.getMessage());
						//write as Fail into status cell in TCModule sheet
						xl.setCellData(TCModule, j, 5, "Fail", outputpath);
						Module_New="False";
					}
					if(Module_Status.equalsIgnoreCase("True"))
					{
						//write as pass into TCsheet in Status cell
						xl.setCellData(TCSheet, i, 3, "Pass", outputpath);
					}
				}
				if(Module_New.equalsIgnoreCase("False"))
				{
					//write as Fail into TCsheet in Status cell
					xl.setCellData(TCSheet, i, 3, "Fail", outputpath);
				}

			}
			else
			{
				//write as blocked for test cases which are flag to N
				xl.setCellData(TCSheet, i, 3, "Blocked", outputpath);
			}

		}
	}

}
