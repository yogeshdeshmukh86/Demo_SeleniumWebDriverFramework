
/*######################################################################################################
'Framework : Selenium Hybrid Framework 
'Author		       	:Yogesh Deshmukh
'Version	    	: 1.0
'Date of Creation	: 21st Nov 2014
'#######################################################################################################
 */

package com.automation;


import static org.junit.Assert.fail;

import java.awt.Robot;
import java.awt.event.InputEvent;
import java.awt.event.KeyEvent;
import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.io.PrintWriter;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Random;
import java.util.Set;
import java.util.concurrent.TimeUnit;

import jxl.Sheet;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.read.biff.BiffException;

import org.apache.commons.io.FileUtils;
import org.apache.commons.lang3.RandomStringUtils;
import org.apache.james.mime4j.field.datetime.DateTime;
import org.junit.After;
import org.junit.Test;
import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.Dimension;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.Point;
import org.openqa.selenium.Proxy;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.UnhandledAlertException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.htmlunit.HtmlUnitDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.interactions.Action;
import org.openqa.selenium.phantomjs.PhantomJSDriver;
import org.openqa.selenium.phantomjs.PhantomJSDriverService;
import org.openqa.selenium.remote.CapabilityType;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.remote.RemoteWebDriver;
import org.openqa.selenium.support.Color;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.subethamail.wiser.Wiser;
import org.junit.Assert;

import javax.mail.Message;
import javax.mail.internet.MimeMessage;


public class SeleniumWebdriverProj<WindowHandler> {

	static BufferedWriter bw = null;
	static BufferedWriter bw1 = null;
	static WebDriver D8;
	Date cur_dt = null;
	String TestSuite;
	String TestScript;
	String ObjectRepository;
	int startrow = 0;
	static String ReportsPath;
	String strResultPath;
	String[] TCNm = null;
	static String exeStatus = "True";
	static String TestReport;
	static int rowcnt;
	static int dtrownum = 1;
	static int ORrowcount = 0;
	// int loopcount = 0;
	static String ORvalname = "";
	static String ORvalue = "";
	static String Action = "";
	static String cCellData = "";
	static String dCellData = "";
	static String fCellData = "";
	static String htmlname = "";
	String[] cCellDataVal = null;
	String[] dCellDataVal = null;
	String ObjectSet;
	String ObjectSetVal = "";
	
	static Sheet DTsheet = null;
	static Sheet ORsheet;
	String Searchtext;
	static int iflag = 0;
	static int screenshotflag = 0;
	static int loopflag = 0;
	static int j = 0;
	static int loopsize = -1;
	int[] loopstart = new int[1];
	int[] loopcount = new int[1];
	int[] loopend = new int[1];
	static int[] loopcnt = new int[1];
	static int[] dtrownumloop = new int[1];
	boolean captureperform = false;
	static boolean capturecheckvalue = false;
	static boolean capturestorevalue = false;
	static Sheet TScsheet;
	static Workbook TScworkbook;
	static int TScrowcount = 0;
	static int loopnum = 1;
	static String TScname;
	static String ActionVal;
	static String alerttext ="";
	String BrowserType;  // Assign with either FF or IE in Config_Framework Excel
	String BrowserTypeText;
	static int DTcolumncount = 0;
	static WebElement elem = null;
	static int FlagIE=0;
	static WebElement dragElement;
	static WebElement dropElement;
	private static Map<String, String> map = new HashMap<String, String>();
	private static Map<String, Float> mapint = new HashMap<String, Float>();
	private Boolean bool_Security;
	public static String failedScreenShot="";	
  //private Wiser wiser;
	/*
	 * This function reads the selenium utility file and identifies where Object
	 * Repository, Test Suite & Test Scripts are located
	 */
 // Email variables
//************************************************************************
	static String Username = "mock account";
	static String Password = "password";
	static String Subject = "Production Testing";
	static String Body = "Production Automation Test Script has Failed & this may be due to Production Server Issues Please Verify manually";
	static String Sender = "deshmukhyog86@gmail.com";
	static String Recepient = "deshmukhyog86@gmail.com";
//*************************************************************************
	
	
	@Test
	public void ReadUtilFile() throws Exception {
		for (int z = 0; z < 1; z++) {
			loopstart[z] = 0;
			loopend[z] = 0;
			loopcnt[z] = 0;
			dtrownumloop[z] = 1;
			loopcount[z] = 0;
		}
		
		{
			//PrintWriter pw = new PrintWriter(new FileWriter("D:/c.html"));
		//pw.println("<TABLE BORDER><TR><TH>Execute<TH>Keyword<TH>Object<TH>Action<Result></TR>");

			Workbook w1 = null;
			try {
				//System.setProperty("http.proxyHost", "132.186.124.49");
				//System.setProperty("http.proxyPort", "8080");
			
				FileInputStream fi = new FileInputStream("D:\\SeleniumKeywordDrivenFramework\\Config_Framework.xls");
			 w1 = Workbook.getWorkbook(fi);
				//FileInputStream fi = new FileInputStream("D:\\Test_Data_Taasera.xls");
				//w1 = Workbook.getWorkbook(new File("D:\\Selenium_Utility.xls"));
				//w1 = Workbook.getWorkbook(new File("D:\\SeleniumKeywordDrivenFramework\\Selenium_Utility.xls"));
				
			} catch (BiffException e) { // TODO Auto-generated catch block
				e.printStackTrace();
			} catch (IOException e) { // TODO Auto-generated catch block
				e.printStackTrace();
			}
			Sheet sheet = w1.getSheet(0);
			TestSuite = sheet.getCell(1, 1).getContents();
			TestScript = sheet.getCell(1, 2).getContents();
			ObjectRepository = sheet.getCell(1, 3).getContents();
			ReportsPath = sheet.getCell(1, 5).getContents();
			TestReport = sheet.getCell(1, 6).getContents();
			BrowserType =sheet.getCell(1, 7).getContents();
			
			System.out.println("The Testsuite is " +TestSuite );
			System.out.println("The Test script is " +TestScript );
			System.out.println("The obj repository is " +ObjectRepository );
			System.out.println("The Report Path is " +ReportsPath );
			System.out.println("The Browser type is " +BrowserType );
			
			//FindExecTestscript(TestSuite, TestScript, ObjectRepository);
		}
		switch (BrowserType.toUpperCase().toString()) {
		
		case "IE":
			FlagIE = 1;
			BrowserTypeText= "Internet Explorer";
			System.out.println("Launching IE and IE Flag is " + FlagIE);
			System.setProperty("webdriver.ie.driver",
					"D://SeleniumKeywordDrivenFramework//IEDriverServer.exe");
			DesiredCapabilities capability = DesiredCapabilities
					.internetExplorer();
			capability
					.setCapability(
							InternetExplorerDriver.INTRODUCE_FLAKINESS_BY_IGNORING_SECURITY_DOMAINS,
							true);
			capability.setCapability("useLegacyInternalServer", true);
			
			capability.setCapability(CapabilityType.ACCEPT_SSL_CERTS,true); 
			
			D8 = new InternetExplorerDriver(capability);
			D8.getWindowHandle();
			D8.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
			D8.manage().window().maximize();
			
			//D8.get("javascript:document.getElementById('overridelink').click();");
			FindExecTestscript(TestSuite, TestScript, ObjectRepository);
			break;
			// For Firefox
		case "FF":
			System.out.println("Launching firefox"  );
			 BrowserTypeText= "FireFox";
			D8 = new FirefoxDriver();
			/*
			 DefaultSelenium ds=new DefaultSelenium("192.168.10.114",8888,"*firefox","");
	          ds.start();
	          ds.windowMaximize();
	          ds.open("/");*/
		
			D8.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
			
			D8.manage().window().maximize();
			FindExecTestscript(TestSuite, TestScript, ObjectRepository);
			break;
		
			
		// for Google chrome we need Chrome driver to be kept at specific path for win7 keep the google driver in
		//	C:\Users\Username\AppData\Local\Google\Chrome\User Data
		case "GC":
		
			
			System.setProperty("webdriver.chrome.driver", "D:\\SeleniumKeywordDrivenFramework\\chromedriver_win32\\chromedriver.exe");
			//System.setProperty("webdriver.chrome.driver", "D:\\SeleniumKeywordDrivenFramework\\Chorme driver\\chromedriver.exe");
			
			BrowserTypeText= "Google Chrome";
			D8 = new ChromeDriver();
			
           //D8.manage().timeouts().implicitlyWait(20, TimeUnit.SECONDS);		
			D8.manage().window().maximize();
			FindExecTestscript(TestSuite, TestScript, ObjectRepository);
			break;
			
		case "HU":
		  
		/* HtmlUnitDriver D8 = new HtmlUnitDriver();
		  
		  
		  ArrayList<String> noProxyHosts = null;
		  D8.setHTTPProxy("http://inbomt003srv.in002.siemens.net/proxy.pac", 8080, noProxyHosts);
		  
		  System.out.println("Started Execution on headless browser"); 
		  // D8 = new HtmlUnitDriver();
		  D8.get("http://www.open2test.org/");
		  
		// Print the title
		System.out.println("Title of the page "+ D8.getTitle());
	    // TODO Auto-generated method stub
*/		  
		  
		  // D8 = new HtmlUnitDriver(); 
		   BrowserTypeText= "HtmlUnit(HeadlessBrowser)";
		   D8 = new HtmlUnitDriver(DesiredCapabilities.chrome()); 
		    
		   ((HtmlUnitDriver) D8).setJavascriptEnabled(true);
		   
		  Proxy proxy = new Proxy(); 
		  proxy.setHttpProxy("127.0.0.1:3128"); 
		  ((HtmlUnitDriver) D8).setProxySettings(proxy); 
		 /* D8.get("http://www.open2test.org/");
		  System.out.println("Title of the page "+ D8.getTitle());*/
		  
		  java.util.logging.Logger.getLogger("com.gargoylesoftware.htmlunit").setLevel(java.util.logging.Level.OFF);
	    java.util.logging.Logger.getLogger("org.apache.http").setLevel(java.util.logging.Level.OFF);
	    
		  FindExecTestscript(TestSuite, TestScript, ObjectRepository);
		  
		
		break;
		
		
		case "PJS":
		  System.out.println("in pj"); 
		  
	       
          DesiredCapabilities caps = new DesiredCapabilities();
	
          
         // caps.setJavascriptEnabled(true);
          caps.setCapability(PhantomJSDriverService.PHANTOMJS_CLI_ARGS, new String[] {"--web-security=no", "--ignore-ssl-errors=yes"});
                
		  
	BrowserTypeText= "Phantom js (HeadlessBrowser)";
	
	File file =new File("D:/SeleniumKeywordDrivenFramework/phantomjs-2.1.1-windows/bin/phantomjs.exe");
	
	System.setProperty("phantomjs.binary.path", file.getAbsolutePath());
	
	D8= new PhantomJSDriver();
	
	D8.navigate().to("https://imp-test.imp-siemens.com/imp-ui/");
  System.out.println("Title of the page "+ D8.getTitle());
	 
  FindExecTestscript(TestSuite, TestScript, ObjectRepository);
  break;
	
			
		default:
			System.out.println("Error : Invalid browser type");	
			
		}
	}

	public void FindExecTestscript(String TestSuite, String TestScript,
			String ObjectRepository) throws Exception {
		try {
			System.out.println("In find executeTestscript" );

			
			int TSrowcount = 0;
			FileInputStream fs = null;
			WorkbookSettings ws = null;
			fs = new FileInputStream(new File(TestSuite));
			ws = new WorkbookSettings();
			ws.setLocale(new Locale("en", "EN"));
			Workbook TSworkbook = Workbook.getWorkbook(fs, ws);
			Sheet TSsheet = TSworkbook.getSheet(0);
			TSrowcount = TSsheet.getRows();
			cur_dt = new Date();
			DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd-HH-mm-ss");
			String strTimeStamp = dateFormat.format(cur_dt);
			String rp = ReportsPath;
			if (rp == "") { // if results path is passed as null, by
				// default 'C:\' drive is used
				rp = "C:\\";
			}

			if (rp.endsWith("\\")) { // checks whether the path ends with
				// '/'
				rp = rp + "\\";
			}
			// TCNm = scriptName.split("\\.");
			strResultPath = rp + "Log" + "/";
			String htmlname1 = rp + "Log" + "/Test_Suite_" + strTimeStamp
					+ ".html";
			File f = new File(strResultPath);
			f.mkdirs();
			bw1 = new BufferedWriter(new FileWriter(htmlname1));
			bw1.write("<HTML><BODY><TABLE BORDER=0 CELLPADDING=3 CELLSPACING=1 WIDTH=100%>");
			bw1.write("<TABLE BORDER=0 BGCOLOR=BLACK CELLPADDING=3 CELLSPACING=1 WIDTH=100%>");
			bw1.write("<TR><TD COLSPAN=7 BGCOLOR=#66699><FONT FACE=VERDANA COLOR=WHITE SIZE=2><B>Testcase Name</B></FONT></TD><TD BGCOLOR=#66699 WIDTH=27%><FONT FACE=VERDANA COLOR=WHITE SIZE=2><B>Status</B></FONT></TD></TR>");
			for (int i = 0; i < TSsheet.getRows(); i++) {
				String TSvalidate = "r";
				if (((TSsheet.getCell(0, i).getContents())
						.equalsIgnoreCase(TSvalidate) == true)) {
					// String TCStatus = "Pass";
					String ScriptName = TSsheet.getCell(1, i).getContents();
					ExecKeywordScript(ScriptName, TestScript, ObjectRepository);
					if (exeStatus.equalsIgnoreCase("Failed")) {
						bw1.write("<TR><TD COLSPAN=7 BGCOLOR=WHITE><FONT FACE=VERDANA COLOR=BLACK SIZE=2><B>"
								+ TCNm[0]
								+ "</B></FONT></TD><TD BGCOLOR=WHITE WIDTH=27%><FONT FACE=VERDANA COLOR=RED SIZE=2><B>"
								+ exeStatus + "</B></FONT></TD></TR>");
					} else {
						bw1.write("<TR><TD COLSPAN=7 BGCOLOR=WHITE><FONT FACE=VERDANA COLOR=BLACK SIZE=2><B>"
								+ TCNm[0]
								+ "</B></FONT></TD><TD BGCOLOR=WHITE WIDTH=27%><FONT FACE=VERDANA COLOR=GREEN SIZE=2><B>"
								+ exeStatus + "</B></FONT></TD></TR>");
					}
				}
				/*
				 * else { System.out.println(TSvalidate); }
				 */
			}
			bw1.close();
		} catch (Exception e) {
			bw.close();
			bw1.close();
		}
	}

	public void ExecKeywordScript(String scriptName, String TestScript,
			String ObjectRepository) throws Exception {

		// Report header
		cur_dt = new Date();
		DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd-HH-mm-ss");
		String strTimeStamp = dateFormat.format(cur_dt);

		if (ReportsPath == "") { // if results path is passed as null, by
									// default 'C:\' drive is used
			ReportsPath = "C:\\";
		}

		if (ReportsPath.endsWith("\\")) { // checks whether the path ends with
											// '/'
			ReportsPath = ReportsPath + "\\";
		}
		TCNm = scriptName.split("\\.");
		strResultPath = ReportsPath + "Log" + "/" + TCNm[0] + "/";
		String htmlname = ReportsPath + "Log" + "/" + TCNm[0] + "/"
				+ strTimeStamp + ".html";
		File f = new File(strResultPath);
		f.mkdirs();
		bw = new BufferedWriter(new FileWriter(htmlname));
		bw.write("<HTML><BODY><TABLE BORDER=0 CELLPADDING=3 CELLSPACING=1 WIDTH=100%>");
		bw.write("<TABLE BORDER=0 BGCOLOR=BLACK CELLPADDING=3 CELLSPACING=1 WIDTH=100%>");
		bw.write("<TR><TD BGCOLOR=#66699 WIDTH=27%><FONT FACE=VERDANA COLOR=WHITE SIZE=2><B>Test Case Name:</B></FONT></TD><TD COLSPAN=7 BGCOLOR=#66699><FONT FACE=VERDANA COLOR=WHITE SIZE=2><B>"
				+ TCNm[0] + "<TD COLSPAN=9 BGCOLOR=#66699><FONT FACE=VERDANA COLOR=WHITE SIZE=2><B>Browser Type: "+ BrowserTypeText +"</B></FONT></TD></FONT></TD></TR>");
		bw.write("<HTML><BODY><TABLE BORDER=1 CELLPADDING=3 CELLSPACING=1 WIDTH=100%>");
		bw.write("<TR COLS=7><TD BGCOLOR=#FFCC99 WIDTH=3%><FONT FACE=VERDANA COLOR=BLACK SIZE=2><B>Row</B></FONT></TD><TD BGCOLOR=#FFCC99 WIDTH=15%><FONT FACE=VERDANA COLOR=BLACK SIZE=2><B>Keyword</B></FONT></TD><TD BGCOLOR=#FFCC99 WIDTH=40%><FONT FACE=VERDANA COLOR=BLACK SIZE=2><B>Description</B></FONT></TD><TD BGCOLOR=#FFCC99 WIDTH=20%><FONT FACE=VERDANA COLOR=BLACK SIZE=2><B>Object</B></FONT></TD><TD BGCOLOR=#FFCC99 WIDTH=20%><FONT FACE=VERDANA COLOR=BLACK SIZE=2><B>Action</B></FONT></TD><TD BGCOLOR=#FFCC99 WIDTH=25%><FONT FACE=VERDANA COLOR=BLACK SIZE=2><B>Execution Time</B></FONT></TD><TD BGCOLOR=#FFCC99 WIDTH=20%><FONT FACE=VERDANA COLOR=BLACK SIZE=2><B>Status</B></FONT></TD></TR>");

		exeStatus = "Pass";
		String scriptPath = TestScript + scriptName;
		System.out.println("The script Path is " +scriptPath );
		
		TScname = scriptName;
		FileInputStream fs1 = null;
		WorkbookSettings ws1 = null;
		// Enter the Username
		//fs1 = new FileInputStream(scriptPath);
		
	// w1 = Workbook.getWorkbook(fi);
		fs1 = new FileInputStream(new File(scriptPath));
		ws1 = new WorkbookSettings();
		ws1.setLocale(new Locale("en", "EN"));
		
		Workbook TScworkbook = Workbook.getWorkbook(fs1, ws1);
		
		TScsheet = TScworkbook.getSheet(0);
		fs1.read();
		TScrowcount = TScsheet.getRows();
		System.out.println("The Row count " +TScrowcount );
		// *This is the Data Table Sheet
		rowcnt = 0;
		// System.out.println("Row count : " + TScrowcount);
		for (j = 0; j < TScrowcount; j++) {
			// Thread.sleep(1000);
			// System.out.println("J : " + j);
			rowcnt = rowcnt + 1;
			String TSvalidate = "r";
			if (((TScsheet.getCell(0, j).getContents())
					.equalsIgnoreCase(TSvalidate) == true)) {
				Action = TScsheet.getCell(1, j).getContents();
				cCellData = TScsheet.getCell(2, j).getContents();
				
				System.out.println("The" + cCellData );
				
				dCellData = TScsheet.getCell(3, j).getContents();
				// Added for visibiliity of Description in Report.
				fCellData= TScsheet.getCell(5, j).getContents();
				String ORPath = ObjectRepository;
				FileInputStream fs2 = null;
				WorkbookSettings ws2 = null;
				try {
					System.out.println("The OR path is"+ORPath);
					
					fs2 = new FileInputStream(new File(ORPath));
					ws2 = new WorkbookSettings();
					ws2.setLocale(new Locale("en", "EN"));
				} catch (Exception e) {
					System.out.println("File not found");
				}
				try {
					Workbook ORworkbook = Workbook.getWorkbook(fs2, ws2);
					ORsheet = ORworkbook.getSheet(0);
					ORrowcount = ORsheet.getRows();
					ActionVal = Action.toLowerCase();
					iflag = 0;

				} catch (Exception e) {
					// System.out.println(e);
					fail("Excel file of not correct.");
				}
				bcellAction(scriptName);
			}// End of Execution

		}// End of If that get all rows in Test Script
		bw.close();
	}// End of For that get all rows in Test Script

	public static void screenshot(int loopn, int rown, String Sname)
			throws Exception {
		try {
			screenshotflag = screenshotflag + 1;
			DateFormat dateFormat = new SimpleDateFormat("yyyyMMdd-HH-mm-ss");
			Date date = new Date();
			String filenamer = "";
			String strTime = dateFormat.format(date);
			Sname = Sname.substring(0, Sname.indexOf("."));
			File screenshot = ((TakesScreenshot) D8)
					.getScreenshotAs(OutputType.FILE);
			TestReport = TestReport.toLowerCase();
			if (TestReport == "")
				TestReport = ReportsPath;
			if (!(TestReport.contains("screen")))
				TestReport = TestReport + "Screenshot/";
			if (loopflag == 0) {
				filenamer = TestReport + Sname + "/" + Sname + "_"
						+ screenshotflag + "_rowno_" + (j + 1) + "_" + strTime
						+ ".png";
			} else {
				filenamer = TestReport + Sname + "/" + Sname + "_"
						+ screenshotflag + "_loop_" + (loopcnt[loopsize] + 1)
						+ "_rowno_" + (j + 1) + "_" + strTime + ".png";
			}
			FileUtils.copyFile(screenshot, new File(filenamer));
			//Added by for failed link  
			failedScreenShot=filenamer;
			System.out.println("failedScreenShot::"+failedScreenShot);
			//Added by for failed link  
		} catch (Exception e) {
			System.out.println(e);
		}
	}

	public static void Update_Report(String Res_type) throws IOException {
		String str_time;
		String[] str_rep = new String[2];
		Date exec_time = new Date();
		DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd-HH-mm-ss");
		str_time = dateFormat.format(exec_time);
		if (Res_type.startsWith("executed")) {
			bw.write("<TR COLS=7><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"
					+ (j + 1)
					+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
					+ Action
					+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=40%><FONT FACE=VERDANA SIZE=2><B>"
					+fCellData
					+ "</B></FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
					+ cCellData
					+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
					+ dCellData
					+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=20%><FONT FACE=VERDANA SIZE=2>"
					+ str_time
					+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = GREEN>"
					+ "Passed" + "</FONT></TD></TR>");
		}
		else if(Res_type.startsWith("alert")) {
			bw.write("<TR COLS=7><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"
					+ (j + 1)
					+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
					+ Action
					+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=40%><FONT FACE=VERDANA SIZE=2><B>"
					+fCellData
					+ "</B></FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
					+ alerttext
					+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
					+ dCellData
					+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=20%><FONT FACE=VERDANA SIZE=2>"
					+ str_time
					+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = ORANGE>"
					+ "Warning" + "</FONT></TD></TR>");
		} 
		
		else if (Res_type.startsWith("failed")) {
			exeStatus = "Failed";
			
			try{
			screenshot(loopnum, TScrowcount, TScname);
			}catch(Exception ex){
				System.out.println("screenshot Exception "+ex);	
				
			}
			System.out
			.println("The STEP FAILED new implementation of hyperlink");	
			bw.write("<TR COLS=7><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"
					+ (j + 1)
					+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
					+ Action
					+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=40%><FONT FACE=VERDANA SIZE=2><B>"
					+fCellData
					+ "</B></FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
					+ cCellData
					+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
					+ dCellData
					+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=20%><FONT FACE=VERDANA SIZE=2>"
					+ str_time
					+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%>"
					+"<a href="+failedScreenShot+">"
					+"<FONT FACE=VERDANA SIZE=2 COLOR = RED>"
					+ "Failed"
					+"</FONT>"
					+"</a>"
					+ "</TD></TR>");
			
		} else if (Res_type.startsWith("loop")) {
			bw.write("<TR COLS=7><th colspan= 7 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=BLUE><div align=left></FONT><FONT FACE=VERDANA SIZE=2 COLOR = BLUE>"
					+ Res_type + "</div></th></FONT></TR>");
		} else if (Res_type.startsWith("missing")) {
			exeStatus = "Failed";
			try{
				screenshot(loopnum, TScrowcount, TScname);
				}catch(Exception ex){
					System.out.println("screenshot Exception "+ex);	
					
				}
			bw.write("<TR COLS=7><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"
					+ (j + 1)
					+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
					+ Action
					+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=40%><FONT FACE=VERDANA SIZE=2><B>"
					+fCellData
					+ "</B></FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
					+ cCellData
					+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
					+ dCellData
					+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=20%><FONT FACE=VERDANA SIZE=2>"
					+ str_time
					+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = ORANGE>"
					+"<a href="+failedScreenShot+">"
					+ "Failed"
					+"</a>"	
					+ "</FONT></TD></TR>");
			bw.write("<TR COLS=7><th colspan= 7 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in Keyword test step number "
					+ (j + 1)
					+ ".Description: The Datatable column name not found</div></th></FONT></TR>");
		} else if (Res_type.startsWith("ObjectLocator")) {
			try{
				screenshot(loopnum, TScrowcount, TScname);
				}catch(Exception ex){
					System.out.println("screenshot Exception "+ex);	
					
				}
			bw.write("<TR COLS=7><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"
					+ (j + 1)
					+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
					+ Action
					+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=40%><FONT FACE=VERDANA SIZE=2><B>"
					+fCellData
					+ "</B></FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
					+ cCellData
					+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
					+ dCellData
					+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=20%><FONT FACE=VERDANA SIZE=2>"
					+ str_time
					+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = ORANGE>"
					+"<a href="+failedScreenShot+">"
					+ "Failed"
					+"</a>"						+ "</FONT></TD></TR>");
			bw.write("<TR COLS=7><th colspan= 7 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in Keyword test step number "
					+ (j + 1)
					+ ".Description: Object Locator is wrong or not supported. Supported Locators are Id,Name,Xpath& CSS</div></th></FONT></TR>");
		}
	}

	public static void Update_Report(String Res_type, Exception e)
			throws IOException {
		String str_time;
		Date exec_time = new Date();
		DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd-HH-mm-ss");
		str_time = dateFormat.format(exec_time);
		exeStatus = "Failed";
		if (Res_type.startsWith("failed")) {
			System.out
			.println("The STEP FAILED");	
			bw.write("<TR COLS=7><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"
					+ (j + 1)
					+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
					+ Action
					+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=40%><FONT FACE=VERDANA SIZE=2><B>"
					+fCellData
					+ "</B></FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
					+ cCellData
					+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
					+ dCellData
					+"<a href="+failedScreenShot+">"
					+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=20%><FONT FACE=VERDANA SIZE=2>"
					+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
					+ "Failed"
					+"</FONT>"
					+"</a>"	
					+ "</TD></TR>");
			bw.write("<TR COLS=7><th colspan=7 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left></FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
					+ e.toString().substring(
							e.toString().indexOf(":") + 1,
							e.toString().indexOf(".",
									e.toString().indexOf(":") + 1) + 1)
					+ "</div></th></FONT></TR>");
		}
	}

	private static void Func_StoreCheck() throws Exception {
		// TODO Auto-generated method stub
		try {
			
			String actval = null;
			String expval;
			Boolean boolval = null;
			String varname;
			String[] cCellDataValCh = cCellData.split(";");
			String ObjectValCh = cCellDataValCh[1];
			String[] dCellDataValCh = dCellData.split(":");
			String ObjectSetCh = dCellDataValCh[0];
			String ObjectSetValCh = "";
			int DTcolumncountCh = 0;
			//DTcolumncountCh = DTsheet.getColumns();
			
			if (dCellDataValCh.length == 2) {
				ObjectSetValCh = dCellDataValCh[1];
			}
			for (int k = 1; k < ORrowcount; k++) {
				String ORName = ORsheet.getCell(0, k).getContents();

				if (((ORsheet.getCell(0, k).getContents())
						.equalsIgnoreCase(ObjectValCh) == true)) {
					String[] ORcellData = new String[3];
					ORcellData = (ORsheet.getCell(3, k).getContents())
							.split("=");
					ORvalname = ORcellData[0];
					///==========================
					System.out
					.println("The object property type is "
							+ ORvalname);	
					//==============================
					//ORvalue = ORcellData[1];

					ORvalue = ORsheet.getCell(3, k).getContents()
							.substring(ORcellData[0].length() + 1);
					//===================================
					System.out
					.println("The object property value is "
							+ ORvalue);	
					//=====================================
					k = ORrowcount + 1;
				}
			}
			if (ObjectSetValCh.contains("dt_")) {
				String ObjectSetValtableheader[] = ObjectSetValCh.split("_");
				int column = 0;
				String Searchtext = ObjectSetValtableheader[1];

				for (column = 0; column < DTcolumncount; column++) {
					if (Searchtext.equalsIgnoreCase(DTsheet.getCell(column, 0)
							.getContents()) == true) {
						ObjectSetValCh = DTsheet.getCell(column, dtrownum)
								.getContents();
						iflag = 1;
					}
				}
				if (iflag == 0) {
					ORvalname = "exit";
				}
			}
			switch (ObjectSetCh.toLowerCase()) {
			case "enabled":
				Func_FindObj(ORvalname, ORvalue);
				boolval = elem.isEnabled();
				actval = boolval.toString();
				break;
				
//Added for Button text comparision 	
			case "text":
			  System.out
        .println("inside case for text for getting text property");
				Func_FindObj(ORvalname, ORvalue);
				actval = elem.getAttribute("value");
				String var2=elem.getText();
				System.out
				.println("Actual value in text of the elem is"
						+ actval);
				break;
//Added  for any link text comparision 
			case "comparetext":
			
				Func_FindObj(ORvalname, ORvalue);
				actval=D8.findElement(By.xpath(ORvalue)).getText();
				System.out
				.println("Actual value in 'comparetext' is "
						+ actval);
				break;	
//Added for getting text value 	
			case "value":
				Func_FindObj(ORvalname, ORvalue);
				actval = new Select(elem).getFirstSelectedOption().getText()
						.toString();
				System.out
				.println("Actual value in 'value' is "
						+ actval);
				break;
			case "visible":
				Func_FindObj(ORvalname, ORvalue);
				try
				{ //D8.manage().timeouts().implicitlyWait(60, TimeUnit.SECONDS);
				boolval = elem.isDisplayed();
				actval = boolval.toString();
				System.out
				.println("Actual value  for visible case is "
						+ actval);
				//elementHighlight(elem);
				}
				catch (NoSuchElementException exception) {
					System.out.println("ELEMENT NOT FOUND");
				//	GmailSender gmailsender = new GmailSender(Username,Password);
		        //	gmailsender.sendMail(Subject, Body, Sender, Recepient, false);
		        	//Thread.sleep(5000);
		        	//System.exit(1);
					Update_Report("failed");
				}
				break;
//			case "matchContent":	
//				Func_FindObj(ORvalname, ORvalue);
//				System.out.println("ORvalue~:::"+ORvalue);
//				System.out.println("elem~:::"+elem);
//
//				break;	
				// Color action is specific to Update Manager (08/07/2014) added by	// Atul/Yogesh
			case "color":
					Func_FindObj(ORvalname, ORvalue);
					String color = elem.getCssValue("color");
					actval = Color.fromString(color).asHex().toUpperCase();
					break;
			// Color action is specific to Update Manager (08/07/2014) added by // Atul/Yogesh
			case "verifyscore":
					System.out.println("inside verifyScore");
					Func_FindObj(ORvalname, ORvalue);
					boolean isvalidColor=false;
					System.out.println("text:"+elem.getText());
					String scorecolor = Color.fromString(elem.getCssValue("color")).asHex().toUpperCase();
					System.out.println("scorecolor:"+scorecolor);
					
					if(Integer.parseInt(elem.getText())<=24)
					{
						if(scorecolor.equalsIgnoreCase("#FF0000"))
						isvalidColor=true;
						
					}else if(Integer.parseInt(elem.getText())>24 && Integer.parseInt(elem.getText())<49)
					{
						if(scorecolor.equalsIgnoreCase("#FFA500"))
						isvalidColor=true;
						
					}else if(Integer.parseInt(elem.getText())<90 && Integer.parseInt(elem.getText())>=49){
						if(scorecolor.equalsIgnoreCase("#C99701"))
						isvalidColor=true;
						
					}else if(Integer.parseInt(elem.getText())>90){
						if(scorecolor.equalsIgnoreCase("#008000"))
						isvalidColor=true;
					}
					actval=Boolean.toString(isvalidColor);
					break;
			case "checked":
				Func_FindObj(ORvalname, ORvalue);
				boolval = elem.isSelected();
				actval = boolval.toString();
				break;
			case "linktext":
				Func_FindObj(ORvalname, ORvalue);
				actval = elem.getText();
				break;
			default:
				actval = "Invalid syntax";
				break;
			}

			if ((ActionVal).equalsIgnoreCase("check")) {
				expval = ObjectSetValCh;
				System.out.println("Inside check : expval:"+expval);
				System.out.println("Inside check : actval:"+actval);
				if (expval.equalsIgnoreCase("On"))
					expval = "True";
				else if (expval.equalsIgnoreCase("Off"))
					expval = "False";
				if (expval.trim().equalsIgnoreCase(actval.trim())) {
					System.out
							.println("Actual value matches with expected value. Actual value is "
									+ actval);
					Update_Report("executed");
				} else {
					System.out
							.println("Actual value doesn't match with expected value. Actual value is "
									+ actval);
					Update_Report("failed");
					if (ORvalname == "exit") {
						Update_Report("missing");
					} else {
						Update_Report("failed");
					}
					if (capturecheckvalue == true) {
						screenshot(loopnum, TScrowcount, TScname);
					}
				}
			} else if ((ActionVal).equalsIgnoreCase("storevalue")) {
				System.out.println("Inside storevalue "+actval);
				varname = ObjectSetValCh;
				if (actval.equalsIgnoreCase("Invalid syntax")) {
					Update_Report("missing");
				} else {
					if (map.containsKey(varname)) {
						map.put(varname, actval);
						Update_Report("executed");
						System.out
								.println("Overwriting the value of the variable "
										+ varname
										+ " to store the value as mentioned in the test case row number"
										+ rowcnt);
						//map.remove(varname);
					} else {
						map.put(varname, actval);
						Update_Report("executed");
						System.out
								.println("Overwriting the value of the variable "
										+ varname
										+ " to store the value as mentioned in the test case row number"
										+ rowcnt);
						if (ORvalname == "exit") {
							Update_Report("missing");
						} else {
						}
					}
				}
				if (capturestorevalue == true) {
					screenshot(loopnum, TScrowcount, TScname);
				}
			}
		} catch (Exception e) {
			Update_Report("failed", e);
			// bw.close();
		}
	}

	@After
	public void close() throws Exception {
		try {
		 // wiser.stop();
			System.out.println("Test Completed.");
			
	 // D8.quit();
		
		} catch (UnhandledAlertException e) {
			System.out.println(e);
			Alert alert1 = D8.switchTo().alert();
			
			alert1.accept();
			
			System.out
			.println("ALERT ACCPETED");
			
			String sNewWindow = D8.getWindowHandle();
			D8.switchTo().window(sNewWindow);
			
			
			System.out
					.println("Because of specification of SeleniumWebDriver, downloading may be failed.");
			System.out
					.println("Please confirm the report file and screenshot about test result.");
		}
	}

	private static void Func_FindObj(String strObjtype, String strObjpath)
			throws  Exception {
		try {
			System.out
			.println("function find object");
			System.out
			.println("The object type is" +strObjtype);
			System.out
			.println("The object property xpath or id is" +strObjpath);
			
			if (strObjtype.length() > 0 && strObjpath.length() > 0) {
				if (strObjtype.equalsIgnoreCase("id")) {
					System.out
					.println("finding object by id");
					elem = D8.findElement(By.id(strObjpath));
				} else if (strObjtype.equalsIgnoreCase("name")) {
					elem = D8.findElement(By.name(strObjpath));
				} else if (strObjtype.equalsIgnoreCase("xpath")) {
					System.out
					.println("finding object by XPATH");
					
					//boolean var1 = D8.findElement(By.xpath(strObjpath)).isDisplayed();				
					
					//if (var1 == true)

					//{
					D8.manage().timeouts().implicitlyWait( 10, TimeUnit.SECONDS );
					
						elem=null;
						boolean unfound = true;
						  int tries = 0;
						  while ( unfound && tries < 10 ) {
						    tries += 1;
						elem = D8.findElement(By.xpath(strObjpath));
						 unfound=false;
						System.out.println("Element Exists value identified by xpath");
						  }
					//}
					
					
				} else if (strObjtype.equalsIgnoreCase("link")) {
					elem = D8.findElement(By.linkText(strObjpath.toString()));
				} else if (strObjtype.equalsIgnoreCase("css")) {
					elem = D8.findElement(By.cssSelector(strObjpath));
				}
			}
		} 
		catch (NoSuchElementException e)
		{
			System.out
			.println("Object not displayed");
			e.printStackTrace();
			Update_Report("failed");
		//	GmailSender gmailsender = new GmailSender(Username,Password);
        	//gmailsender.sendMail(Subject, Body, Sender, Recepient, false);
			//e.getStackTrace();
			elem = null;
			//System.exit(1);
			//return;
			
		}
		catch (Exception e) {
			System.out
			.println("Not able to find the object");
			Update_Report("failed",e);
		System.out.println(e.toString());
			elem = null;
			
		}
	}

	public static int ifConditionSkipper(String strifConditionStatus)
			throws Exception {
		try {
			System.out.println("Inside condition skipper");
			String strKeyword;
			int intLogicStartRow, intLogicEndRow, intIfEndConditionCount, intIfConditionCount;
			String strKeyWord;
			intIfConditionCount = 1;
			intIfEndConditionCount = 0;
			
			System.out.println("The String if condtion status is " +strifConditionStatus );
			
			if (strifConditionStatus.equalsIgnoreCase("false")) {
				intLogicStartRow = j;
				System.out.println("Start Logic Row is"+intLogicStartRow);
				do {
					System.out.println("IN WHILE loop of condition skipper");
					j = j + 1;
					System.out.println("The value of Rowcount  is" +j);
				
					strKeyword = TScsheet.getCell(1,j).getContents();
					System.out.println(strKeyword);
					
					if (strKeyword.equalsIgnoreCase("Condition")) {
						System.out.println(strKeyword);
						intIfConditionCount = intIfConditionCount + 1;
					}
					if (strKeyword.equalsIgnoreCase("Endcondition")) {
						intIfEndConditionCount = intIfEndConditionCount + 1;
						if (intIfConditionCount == intIfEndConditionCount) {
							j = j + 1;
							break;
						}
					}

				} while (true);
			}
		} catch (Exception e) {
			e.printStackTrace();
//			System.out.println("The Exception is " +e.getMessage());

		}
		return j;
	}

	public String Func_IfCondition(String strConditionArgs) throws Exception {
		int iFlag = 1;
		String str[] = strConditionArgs.split(";");
		String value1 = str[0];
		String value2 = str[2];
		String strOperation = str[1];
		strOperation = strOperation.toLowerCase().trim();

		switch (strOperation.toLowerCase()) {
		case "equals":
			if (value1.substring(0, 1).equalsIgnoreCase("#")) {
				value1 = map.get(value1.substring(1, (value1.length())));
				System.out
						.println("Variable used in condition statement has value: "
								+ value1);
			    if (value1==null) {
					iFlag = 1;
					System.out
					.println("Null value for Variable");
				}
			    else if (value1.trim().equalsIgnoreCase(value2.trim())) {
					iFlag = 0;
				}

			} else if (value1.trim().equalsIgnoreCase(value2.trim())) {
				iFlag = 0;
			}
			
			break;

		case "notequals":
			if (value1.substring(0, 1).equalsIgnoreCase("#")) {
				value1 = map.get(value1.substring(1, (value1.length())));
				System.out
						.println("Variable used in condition statement has values: "
								+ value1);
				if (value1.trim().equalsIgnoreCase(value2.trim())) {
					iFlag = 0;
				}
			} else if (!value1.trim().equals(value2.trim())) {
				iFlag = 0;
			}
			break;

		case "greaterthan":
			if (isInteger(value1) && isInteger(value2)) {
				if (Integer.parseInt(value1) > Integer.parseInt(value2)) {
					iFlag = 0;
				}
			} else {
				// Report
			}
			break;
		case "lessthan":
			if (isInteger(value1) && isInteger(value2)) {
				if (Integer.parseInt(value1) < Integer.parseInt(value2)) {
					iFlag = 0;
				}
			} else {
				// Report
			}
			break;
		default:
			Update_Report("missing");

		}
		
		if (iFlag == 0) {
			System.out
			.println("Returning TRUE for if  condition");
			return "true";
		} else {
			System.out
			.println("Returning FALSE for if  condition");
			return "false";
		}

	}

	public void arrayresize() {
		if (loopstart.length <= loopsize) {
			loopstart = Arrays.copyOf(loopstart, loopstart.length + 1);
			loopend = Arrays.copyOf(loopend, loopend.length + 1);
			loopcnt = Arrays.copyOf(loopcnt, loopcnt.length + 1);
			dtrownumloop = Arrays.copyOf(dtrownumloop, dtrownumloop.length + 1);
			loopcount = Arrays.copyOf(loopcount, loopcount.length + 1);
		}
	}

	public void bcellAction(String scriptName) throws Exception {
		try {
			switch (ActionVal.toLowerCase()) {

			case "loop":
				startrow = j;
				dtrownum = 1;
				loopsize = loopsize + 1;
				if (loopsize >= 1) {
					arrayresize();
				}
				loopflag = 1;
				loopcount[loopsize] = Integer.parseInt(cCellData);
				loopstart[loopsize] = j;
				loopcnt[loopsize] = 0;
				dtrownumloop[loopsize] = dtrownum;
				Update_Report("loop : " + "Start of loop : " + (loopsize+1));
				Update_Report("executed");
				break;
			case "endloop":
				loopcnt[loopsize] = loopcnt[loopsize] + 1;
				loopnum = loopnum + 1;
				if (loopcnt[loopsize] == loopcount[loopsize]) {
					Update_Report("loop" + " End of Loop : " + (loopsize + 1)
							+ " : Loop count : " + loopcnt[loopsize]);
					loopsize = loopsize - 1;
					if (loopsize >= 0)
						dtrownum = dtrownumloop[loopsize];
					else {
						dtrownum = 1;
						loopflag = 0;
					}
					Update_Report("executed");
				} else {
					j = loopstart[loopsize];
					dtrownum = dtrownum + 1;
					dtrownumloop[loopsize] = dtrownum;
					Update_Report("loop" + " End of Loop : " + (loopsize + 1)
							+ " : Loop count : " + loopcnt[loopsize]);
				}
				rowcnt = 1;
				break;
				
				//Email Verification:
			case"verifyemail":
			  System.out.println("Inside Email");
			 Wiser wiser = new Wiser();
			  wiser.setPort(5025);
			  wiser.setHostname("localhost");
			  wiser.start();
			 wiser.accept("imp.noreply@gmail.com", "testuserdef@yopmail.com");
			Assert.assertNotNull(wiser.getMessages().size());
			  //Assert.assertTrue(wiser.getMessages()).hasSize(1);
			  MimeMessage message = wiser.getMessages().iterator().next().getMimeMessage();
			  //Assert.assertTrue(message.getSubject()).isEqualTo("Here is a sample subject !");
			  
			 Assert.assertEquals("Email validation Failed", "Service Offering accepted", message.getSubject());
			
				//For iteration through table rows & columns of a table
			
			 
			 
				// For calender specific to imp project not generic as of now
			
			case"tofromdate":
			  Calendar now = Calendar.getInstance();
			  int day = now.get(Calendar.DAY_OF_MONTH);
			  String Day=String.valueOf(day);
        System.out.println("In To from date" +(day));
        
       System.out.println(cur_dt.getTime());
        D8.findElement(By.xpath(".//*[@id='mainView']/div/div/div[8]/table/tbody/tr/td/div/table/tbody/tr/td/div/div/form/div[1]/div/div[1]/div[4]/div[5]/div/div/datepicker/input")).click();
        selectJQueryDate(Day);
        D8.findElement(By.xpath("html/body/div[5]/div/table/tbody/tr[2]/td[2]/div/table/tbody/tr/td/div/div/div/div/form/div[1]/div/div[2]/div[6]/div/div[2]/input")).click();
        D8.findElement(By.xpath(" html/body/div[6]/div/table/tbody/tr[1]/td/table/tbody/tr/td[3]/div")).click();
        Thread.sleep(5000);
        if(Day=="31")
        {
        selectJQueryDate("28");
        }
        else
        {
          selectJQueryDate(Day);
        }
          
        break;
        
			case"tableiterate":
			  boolean breakIt = true;
        while (true) {
        breakIt = true;
        try {
           
          TableSearch();
          
          
        } catch (Exception e) {
            if (e.getMessage().contains("element is not attached")) {
                breakIt = false;
            }
        }
        if (breakIt) {
            break;
        }

    }
        Update_Report("executed");
        break;
        
//Table Iterate Ends here     
		/*//	case "launchapp_basicauth":
				
				System.out.println("Launching url with basic auth");
				
			//	String URL = "http:// + username + ":" + password + "@" + "link";
					    
						D8.get("http://username:password@www.domain.com");
					    D8.manage().window().maximize(); 
				
			//	break;
				
			//case "switchtab":
				Set<String> Pages = D8.getWindowHandles();
				Object WebPages[] = Pages.toArray();
				D8.switchTo().window(WebPages[1].toString());
			//	break;
*/                                                                                                                                                                                                                            
       
			case "launchapp":
				
				D8.get(cCellData);
			System.out.println("Title of the page "+ D8.getTitle());
				
				//D8.navigate().to(cCellData);
				//Thread.sleep(5000);
				// #region SSL workaround for IE
			/*	
			
				D8.get("javascript:document.getElementById('overridelink').click();");
			        	//D8.navigate().to("javascript:document.getElementById==('overridelink').click()");
			        
				D8.findElement(By.xpath("//*[@id='overridelink']")).click();
				*/
			
				//D8.navigate().to("javascript:document.getElementById('overridelink').click()");
				//handling security certificate
				/* if (FlagIE == 1);
				 {
					 try
					 {
		          bool_Security = D8.findElement(By.id("overridelink")).isDisplayed();
		         
		          
					  if (bool_Security = true);
					  {
						  D8.findElement(By.id("overridelink")).click();
					  }
					//D8.navigate().to("javascript:document.getElementById('overridelink').click()");	
					 }
					 catch (NoSuchElementException e)
						{
							System.out
							.println("OVERIDE SECUIRTY LINK FOR IE Doesnot Exists ");
							//Update_Report("failed");
							//e.getStackTrace();
							elem = null;
							
							//return;
						}
				 }*/
				Update_Report("executed");
				break;
			case "wait":
				int Timeout = Integer.parseInt(cCellData);
				D8.manage().timeouts().implicitlyWait(Timeout, TimeUnit.MINUTES);
				//Thread.sleep(Long.parseLong(cCellData));
				Update_Report("executed");
				break;
			case "threadwait":
				Thread.sleep(Long.parseLong(cCellData));
				
				break;
		   case "condition":
				System.out.println("IN CONDITION");
				
				String strConditionStatus = Func_IfCondition(cCellData);
				
				System.out.println(strConditionStatus);
				
				if (strConditionStatus.equalsIgnoreCase("false")) {
					System.out.println("THE STATUS OF Condition is" +strConditionStatus);
					j = ifConditionSkipper(strConditionStatus);
					
					j = j + 1;
					System.out.println(j);
				}
				Update_Report("executed");
				break;

			case "endcondition":
				Update_Report("executed");
				break;

			case "screencaptureoption":
				String[] sco = cCellData.split(";");

				for (int s = 0; s < sco.length; s++) {
					if (sco[s].equalsIgnoreCase("perform")) {
						captureperform = true;
					}
					if (sco[s].equalsIgnoreCase("storevalue")) {
						capturestorevalue = true;
					}
					if (sco[s].equalsIgnoreCase("check")) {
						capturecheckvalue = true;
					}

				}
				Update_Report("executed");
				break;
			case "importdata":
				System.out
				.println("In Case IMPORT DATA");
				// Runtime rt = Runtime.getRuntime();
				// Process p = rt.exec("D://AutoITScript/FileDown.exe");
				String xcelpath = cCellData;
				FileInputStream fs3 = null;
				WorkbookSettings ws3 = null;
				System.out
				.println("the Testdata Excel is"+xcelpath);
				
				fs3 = new FileInputStream(new File(xcelpath));
				ws3 = new WorkbookSettings();
				ws3.setLocale(new Locale("en", "EN"));
				Workbook DTworkbook = Workbook.getWorkbook(fs3, ws3);
				DTsheet = DTworkbook.getSheet(0);
				int DTrowcount = DTsheet.getRows();
				String DTName;
				Update_Report("executed");
				break;

			case "screencapture":
				screenshot(loopnum, TScrowcount, TScname);
				Update_Report("executed");
				break;
			case "check":
				Func_StoreCheck();
				break;
			case "storevalue":
				Func_StoreCheck();
				break;
			case "alert":
             Alert alert1 = D8.switchTo().alert();
             
				alerttext=  alert1.getText();
					
				alert1.accept();
				
				System.out
				.println("ALERT ACCPETED");
				Thread.sleep(2000);
				String sNewWindow = D8.getWindowHandle();
				Thread.sleep(2000);
				Update_Report("alert");
				D8.switchTo().window(sNewWindow);
				alerttext=null;
				break;
		//Addded by Atul Parihar for Enter  to Download  List in UM portal // 30/09/2014	
			case "enter":	
				elem.sendKeys(Keys.ENTER);
				break;	
				
			case "quit":	
				System.out.println("INSIDE QUIT ACTION");
				D8.quit();
				break;	
	/*		case "DownloadList":	
				//Load Dll file 	
				System.setProperty("jacob.dll.path", "E:\\SeleniumKeywordDrivenFramework\\jacob-1.17-M2-x86.dll");
				LibraryLoader.loadJacobLibrary(); 
				
				WindowHandler handler = new WindowHandler();
				WindowElement runWindowElement = handler.getWindowElement("Opening SHADOW.BL");
				WindowElement okButton = handler.findElementByName(runWindowElement,"OK");

				handler.click(okBut	
				
				
				
				
				break;	*/
				
			    
					case "mousemove":
						 System.out.println("Mouse Movement");
						 
					elem=D8.findElement(By.xpath("/html/body/app-root/mat-sidenav-container/mat-sidenav/button[1]"));
					
					int xaxis = elem.getLocation().x;
					 
					int yaxis=elem.getLocation().y;
					 
					int width = elem.getSize().width;
					int height= elem.getSize().height;
					
					 System.out.println("Mouse postion" +xaxis  +yaxis  +width  +height );
					Robot robot = new Robot();
					
					robot.mouseMove(xaxis+width/2+180, yaxis+height/2);
					break;	
				
				
				
				
			case "perform":
				System.out.println("In Case perform");
				System.out
				.println("The data is"+cCellData.toString());
				cCellDataVal = cCellData.split(";");
				dCellDataVal = null;
				String ObjectVal = cCellData.substring(
						cCellDataVal[0].length() + 1, cCellData.length());
				dCellData.toString();
				
				if (dCellData.contains(":")) {
					dCellDataVal = dCellData.split(":");
					ObjectSet = dCellDataVal[0].toLowerCase();
					ObjectSetVal = dCellDataVal[1];
					System.out
					.println("The value selected from dropdown is " + ObjectSetVal);
				} else {
					ObjectSet = dCellData.toString();
				}
				DTcolumncount = 0;
		// Snippet For adding/setting Random Number to a text box use set:Random_integer in Test script
				if  (ObjectSetVal.contains("Random_integer")){
					 Random rand = new Random();
					 int num = rand.nextInt((14030 - 100) + 5) + 11302;
					 ObjectSetVal = String.valueOf(num);
					 System.out
						.println("The Random Number entered is" + ObjectSetVal);
				}
				
			// Snippet For adding/setting Random service alias & service name imp specific keywords
			  if  (ObjectSetVal.contains("Random_servicealias"))
			  {
			    StringBuilder servicealias = new StringBuilder("Test ");
			    String randomservice= RandomStringUtils.randomAlphabetic(3);
			    ObjectSetVal=servicealias.append(randomservice).toString();
			 }
			  
			  if  (ObjectSetVal.contains("Random_servicename"))
        {
          StringBuilder servicealias = new StringBuilder(" Sharing");
          String randomservice= RandomStringUtils.randomAlphabetic(4);
          StringBuilder strservice = new StringBuilder(randomservice);
          ObjectSetVal=strservice.append(servicealias).toString();
       }
			  
				
				/*
				if  (ObjectSetVal.contains("Alert")){
					Alert alert1 = D8.switchTo().alert();
					
					alert1.accept();
					System.out
					.println("ALERT ACCEPTED");
					 
				}*/
				
				
			//================================================================================		 
			    
				// Changes made in order to work with open office for MS office make necessary changes
				if (ObjectSetVal.contains("dt_"))
					DTcolumncount = DTsheet.getColumns();
				for (int k = 1; k < ORrowcount; k++) {
					String ORName = ORsheet.getCell(0, k).getContents();
					System.out
					.println("The or name is"+ORName);
					
					System.out
					.println("The obj val is "+ObjectVal);
					
					
					if (((ORsheet.getCell(0, k).getContents())
							.equalsIgnoreCase(ObjectVal) == true)) {
						String[] ORcellData = new String[3];
						ORcellData = (ORsheet.getCell(3, k).getContents())
								.split("=");
						ORvalname = ORcellData[0]; // OR  VALUE NAME
						
						System.out
						.println("The or value name is"+ORvalname);
						
						ORvalue = ORsheet.getCell(3, k).getContents()
								.substring(ORcellData[0].length() + 1);
						
						System.out
						.println("The or value name is"+ORvalue);
						
						
						k = ORrowcount + 1;
					}
				}
				dCellAction();
				
				/*if (ObjectSetVal.contains("dt_"))
				DTcolumncount = DTsheet.getColumns();
			for (int k = 0; k < ORrowcount; k++) {
				String ORName = ORsheet.getCell(1, k).getContents();

				if (((ORsheet.getCell(1, k).getContents())
						.equalsIgnoreCase(ObjectVal) == true)) {
					String[] ORcellData = new String[3];
					ORcellData = (ORsheet.getCell(4, k).getContents())
							.split("=");
					ORvalname = ORcellData[0]; // OR  VALUE NAME
					
					System.out
					.println("The or value name is"+ORvalname);
					
					ORvalue = ORsheet.getCell(4, k).getContents()
							.substring(ORcellData[0].length() + 1);
					
					System.out
					.println("The or value name is"+ORvalue);
					
					
					k = ORrowcount + 1;
				}
			}	*/
			
				
				
				// End of Objectset Switch
				// Update_Report("executed");
			}// End of Actval Switch
		} /*
		 * catch (UnhandledAlertException e) {
		 * 
		 * for (String id : D8.getWindowHandles()) {
		 * System.out.println("WindowHandle: " + D8.switchTo().window(id));
		 * 
		 * } System.out.println(e); System.out .println(
		 * "Because of specification of SeleniumWebDriver, downloading may be failed."
		 * ); System.out .println(
		 * "Please confirm the report file and screenshot about test result.");
		 * }
		 * 
		
		 */
		catch (NullPointerException ex)
		{
			Update_Report("failed", ex);
			System.out.println("Null pointer Exception");
		}
		
		catch (Exception ex) {
			Update_Report("failed", ex);
			//Update_Report("failed");
			System.out.println(ex);
			String Error= ex.getMessage();
			
			System.out.println("Error Message"  +Error );
			if (Error.contains("Modal dialog present: Selected option cannot be added in combination with rawbytes"))	
			{
				Alert alert1 = D8.switchTo().alert();
				
				alert1.accept();
				
				System.out
				.println("ALERT ACCPETED");
				
				String sNewWindow = D8.getWindowHandle();
				D8.switchTo().window(sNewWindow);
				
				
			}
			
			System.out.println("------Error Information :SELENIUM KDF------");
			System.out.println("Current Script:" + scriptName);
			System.out.println("Current ScriptPath:" + TestScript);
			System.out
					.println("Using ObjectRepositoryPath:" + ObjectRepository);
			System.out.println("Current Keyword:" + Action);
			System.out.println("Current ObjectDetails:" + cCellData);
			System.out.println("Current ObjectDetailsPath:" + ORvalue);
			System.out.println("Current Action:" + dCellData);
			System.out.println("------ERROR INFORMATION------");
			//fail("Cannot test normally by FRAMEWORK.");
			 return;
			 
		}
	}

	public void readAttributeforPerform() throws Exception {
		try {
			System.out.println("In read Attribute for Perform");
			if (ObjectSetVal.length() > 0) {
				if (ObjectSetVal.substring(0, 1).equalsIgnoreCase("#")) {
				  
				  System.out.println("In if for # store variable");
				  
					ObjectSetVal = map.get(ObjectSetVal.substring(1,
							(ObjectSetVal.length())));
				} else if (ObjectSetVal.contains("dt_")) {
					System.out.println("The dt_ is present");
					String ObjectSetValtableheader[] = ObjectSetVal.split("_");
					int column = 0;
					String Searchtext = ObjectSetValtableheader[1];
					for (column = 0; column < DTcolumncount; column++) {
						if (Searchtext.equalsIgnoreCase(DTsheet.getCell(column,
								0).getContents()) == true) {
							ObjectSetVal = DTsheet.getCell(column, dtrownum)
									.getContents();
							System.out.println("The Data is" + ObjectSetVal);
							iflag = 1;
						}

					}
					if (iflag == 0) {
						ORvalname = "exit";
					}
				}
			}
		} catch (Exception e) {
			Update_Report("failed", e);
		}

	}

	public void dCellAction() throws Exception {
		try {
			System.out.println("In d cell action");
			readAttributeforPerform();
			
			System.out.println("the or value passed to find object func is " + ORvalue);
			
			Func_FindObj(ORvalname, ORvalue);
			if (elem == null) {
				return;
			} else {
				switch (ObjectSet.toLowerCase()) {
				case "set":
					elem.clear();
					System.out.println("In set Keyword");
				//	elementHighlight(elem);
					elem.sendKeys(ObjectSetVal);
					
					/*StringBuffer inputvalue = new StringBuffer();
             inputvalue.append(ObjectSetVal);
					((RemoteWebDriver) D8).executeScript(
							"arguments[0].value=arguments[0].value + '"
									+ inputvalue.toString() + "';", elem);
					*/
					
					
					Update_Report("executed");
					if (captureperform == true) {
						screenshot(loopnum, TScrowcount, TScname);
					}
					break;
				case "listselect":

					Func_FindObj(ORvalname, ORvalue);
					String[] listvalues = ObjectSetVal.split(",");
					List<WebElement> listboxitems = elem.findElements(By
							.tagName("option"));
					Select chooseoptn = new Select(elem);
					chooseoptn.deselectAll();
					for (WebElement opt : listboxitems) { // System.out.println(opt.getText());
						for (int i = 0; i < listvalues.length; i++) {
							if (opt.getText().equalsIgnoreCase(listvalues[i])) {
								// System.out.println(listvalues[i]);
								chooseoptn.selectByVisibleText(opt.getText());
							}
						}
					}
					Update_Report("executed");
					if (captureperform == true) {
						screenshot(loopnum, TScrowcount, TScname);
					}

					break;

				case "select":
					// readAttributeforPerform();
					// Func_FindObj(ORvalname, ORvalue);
					try
					{
					
					new Select(elem).selectByVisibleText(ObjectSetVal);
					Update_Report("executed");
					if (captureperform == true) {
						screenshot(loopnum, TScrowcount, TScname);
					}

					break;
					}
					 catch (UnhandledAlertException e) {
							System.out.println(e);
							Alert alert1 = D8.switchTo().alert();
							
							alert1.accept();
							
							System.out
							.println("ALERT ACCPETED");
							
							String sNewWindow = D8.getWindowHandle();
							D8.switchTo().window(sNewWindow);
					 }

				case "check":
					// readAttributeforPerform();
					// Func_FindObj(ORvalname, ORvalue);
				  
				  System.out.println(" In check box case");
					if (elem.isSelected()
							&& dCellDataVal[1].equalsIgnoreCase("On")) {

					} else if (elem.isSelected()
							|| dCellDataVal[1].equalsIgnoreCase("On")) {
					  System.out.println(" In check box case in else if");
						elem.click();
					}
					Update_Report("executed");
					if (captureperform == true) {
						screenshot(loopnum, TScrowcount, TScname);
					}
					break;

				case "click":
					// Func_FindObj(ORvalname, ORvalue);
				// elementHighlight(elem);
				  //((JavascriptExecutor)D8).executeScript("arguments[0].click();",elem);
			    elem.click();
				  System.out.println("Clicked on element" + elem);
					Update_Report("executed");
					if (captureperform == true) {
						screenshot(loopnum, TScrowcount, TScname);
					}
					break;
					
					 
					
	//Addded for upload rule files in UM portal // 07/08/2014				
				case "fileupload":
					//Func_FindObj(ORvalname, ORvalue);
					elem.sendKeys(ObjectSetVal.toString()); 
					Update_Report("executed");
					if (captureperform == true) {
						screenshot(loopnum, TScrowcount, TScname);
					}
					break;	
					
	//Addded for Right click functionality in UM portal // 06/08/2014				
				case "rightclick":
					//Func_FindObj(ORvalname, ORvalue);
					Actions oAction3 = new Actions(D8);
					Thread.sleep(3000);
					oAction3.moveToElement(elem);
					oAction3.contextClick(elem).perform();// Right Click
					Thread.sleep(3000);
					Update_Report("executed");
  				    break;	
					
//Addded to  serach pattern functionality in UM portal // 18/08/2014				 
				case "enter":	
					elem.sendKeys(Keys.ENTER);
					break;
//Addded for Selecting down from weblist functionality in dropdown // 18/08/2014	
				case "down":	
					elem.sendKeys(Keys.DOWN);
					break;
//Addded for wait until Object is available to perform any task 22/08/2014	
				case "waituntil":
					System.out.println("inside waituntil");
					WebDriverWait wait = new WebDriverWait(D8, 1000);
					if(ORvalname.equals("xpath")){
					    wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(ORvalue)));
					}
					if(ORvalname.equals("id")){
					    wait.until(ExpectedConditions.visibilityOfElementLocated(By.id(ORvalue)));
					}
					Update_Report("executed");
					if (captureperform == true) {
						screenshot(loopnum, TScrowcount, TScname);
					}
				break;
//-------------Addded to perform Double Click feature------------
				case "doubleclick":
					//Func_FindObj(ORvalname, ORvalue);
					Actions action4 = new Actions(D8);
					Thread.sleep(3000);
					action4.moveToElement(elem).doubleClick(elem).build().perform();
					Thread.sleep(3000);
					Update_Report("executed");
  				    break;		
  				    
  		
//-------------Addded  for Drag and Drop feature------
				case "drag":
																
					if(ORvalname.equals("xpath")){
						System.out.println("In drag now");
					 dragElement=D8.findElement(By.xpath(ORvalue));  
					}
					if(ORvalname.equals("id")){
					dragElement=D8.findElement(By.id(ORvalue));  
					}
				
				break;
				case "drop":
					
				
					Actions builder = new Actions(D8);  // Configure the Action  
					  System.out.println(" Dropping  now");
					  if(ORvalname.equals("xpath")){
						  System.out.println(" Dropping  now with xpath");
						  dropElement=D8.findElement(By.xpath(ORvalue));  
							}
					if(ORvalname.equals("id")){
								dropElement=D8.findElement(By.id(ORvalue));  
							}
					
				
					System.out.println("dragElement::"+dragElement);
					System.out.println("dropElement::"+dropElement);
					
			//	int xdrop=	dragElement.getLocation().x+50;
				//	dragAndDropElement(dragElement,dropElement,xdrop ); 
					
				//	Drag_Drop();  //This is jquery approach
					
				//	Fundragdrop(dragElement,dropElement); This is another jqeuery approach
					
					 // Action dragAndDrop = builder.clickAndHold(dragElement).moveToElement(dropElement,2,2).release(dropElement).build();  // Get the action 
					    int xto=	dropElement.getLocation().x;
						int yto=	dropElement.getLocation().y;
						
						System.out.println(" The Drop Element X & Y Cord are   "   +xto +yto);
					  builder.clickAndHold(dragElement).moveToElement(dropElement).pause(3000).perform();  // Get the action  
					  
					  builder.clickAndHold(dragElement).moveByOffset(xto,yto).build().perform();;
					  Thread.sleep(3000);// add 2 sec wait
					
					  builder.release(dropElement).build().perform();
					//  dragAndDrop.perform(); // Execute the Action  
					    System.out.println("Complete the  drop action");
				break;
				
				
				
//-----------------------Drag and Drop Ends--------------------------
				
				
				
				}
			}
		} catch (Exception e) {
		  e.printStackTrace();
			Update_Report("failed", e);
		}
	}
	
	
	public void Drag_Drop() throws IOException /// This will work on downgraded version of chrome
	
	{
		System.out.println("INSIDE DRAG DROP");
		
	    String js_filepath = "D:\\SeleniumWebdriverProj\\Resources\\drag_and_drop_helper.js";
	    String java_script="";
	    String text;

	    BufferedReader input = null;
		try {
			input = new BufferedReader(new FileReader(js_filepath));
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	    StringBuffer buffer = new StringBuffer();

	    while ((text = input.readLine()) != null)
	        buffer.append(text + " ");
	        java_script = buffer.toString();

	    input.close();
/*
	    WebElement source = dragElement;
	    WebElement target = dropElement;*/
	    String source= "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/div[2]/gridster/div/gridster-item[4]/div/div[1]/app-average-speed/div/div[1]/div/div[1]]";
	    String target= "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/div[2]/gridster/div/gridster-item[2]/div/div[1]/app-parking-districts/div/div[1]/div/div[1]]";
	   
	    java_script = java_script+"$('#"+source+"').simulateDragDrop( '#" +target+ "');" ;
	    ((JavascriptExecutor)D8).executeScript(java_script);
		
	}
	
	
	
	public void dragAndDropElement(WebElement dragFrom, WebElement dragTo, int   
			xOffset) throws Exception { 
		
		System.out.println("Implementation wiht Robot Class");
			         //Setup robot 
			         Robot robot = new Robot(); 
			         robot.setAutoDelay(50); 

			         //Fullscreen page so selenium coordinates work 
			         //robot.keyPress(KeyEvent.VK_F11); 
			         Thread.sleep(5000); 

			         //Get size of elements 
			         Dimension fromSize = dragFrom.getSize(); 
			         Dimension toSize = dragTo.getSize(); 

			         //Get centre distance 
			         int xCentreFrom = fromSize.width / 2; 
			         int yCentreFrom = fromSize.height / 2; 
			         int xCentreTo = toSize.width / 2; 
			         int yCentreTo = toSize.height / 2; 

			         //Get x and y of WebElement to drag to 
			         Point toLocation = dragTo.getLocation(); 
			         Point fromLocation = dragFrom.getLocation(); 
           
			         //Make Mouse coordinate centre of element 
			         toLocation.x += xOffset + xCentreTo; 
			         toLocation.y += yCentreTo+50; 
			         fromLocation.x += xCentreFrom; 
			         fromLocation.y += yCentreFrom; 
			         System.out.println(" The X cord of drop elem is"  +fromLocation.x  +"The Y cord of drop elem is" +fromLocation.y );
			         
			         System.out.println(" The X cord of drag elem is"  +toLocation.x  +"The Y cord of drop elem is" +toLocation.y );
			         
			         //Move mouse to drag from location 
			         robot.mouseMove(fromLocation.x+10, fromLocation.y); 

			         //Click and drag 
			         robot.mousePress(InputEvent.BUTTON1_MASK); 

			         //Drag events require more than one movement to register 
			         //Just appearing at destination doesn't work so move halfway first 
			         robot.mouseMove(((toLocation.x - fromLocation.x) / 2) +   
			fromLocation.x, ((toLocation.y - fromLocation.y) / 2) + fromLocation.y); 
			         
			         robot.mouseMove(((toLocation.x - fromLocation.x) / 2) +   
			     			fromLocation.x, ((toLocation.y - fromLocation.y) / 2) + fromLocation.y); 

			         //Move to final position 
			         robot.mouseMove(toLocation.x, toLocation.y); 
			         robot.mouseMove(toLocation.x, toLocation.y);
			         robot.mouseMove(toLocation.x, toLocation.y);

			         //Drop 
			         robot.mouseRelease(InputEvent.BUTTON1_MASK); 
	
	}
	
	
	public void Fundragdrop(WebElement dragElement2, WebElement dropElement2) {
	   /* WebElement LocatorFrom = D8.findElement(dragElement2);
	    WebElement LocatorTo = D8.findElement(dropElement2);*/
	   /* String xto=Integer.toString(LocatorTo.getLocation().x);
	    String yto=Integer.toString(LocatorTo.getLocation().y);*/
	    
		int xto=	dropElement.getLocation().x;
		int yto=	dropElement.getLocation().y;
		
		System.out.println(" The Drop Element X & Y Cord are   "   +xto +yto);
		
	    ((JavascriptExecutor)D8).executeScript("function simulate(f,c,d,e){var b,a=null;for(b in eventMatchers)if(eventMatchers[b].test(c)){a=b;break}if(!a)return!1;document.createEvent?(b=document.createEvent(a),a==\"HTMLEvents\"?b.initEvent(c,!0,!0):b.initMouseEvent(c,!0,!0,document.defaultView,0,d,e,d,e,!1,!1,!1,!1,0,null),f.dispatchEvent(b)):(a=document.createEventObject(),a.detail=0,a.screenX=d,a.screenY=e,a.clientX=d,a.clientY=e,a.ctrlKey=!1,a.altKey=!1,a.shiftKey=!1,a.metaKey=!1,a.button=1,f.fireEvent(\"on\"+c,a));return!0} var eventMatchers={HTMLEvents:/^(?:load|unload|abort|error|select|change|submit|reset|focus|blur|resize|scroll)$/,MouseEvents:/^(?:click|dblclick|mouse(?:down|up|over|move|out))$/}; " +
	    "simulate(arguments[0],\"mousedown\",0,0); simulate(arguments[0],\"mousemove\",arguments[1],arguments[2]); simulate(arguments[0],\"mouseup\",arguments[1],arguments[2]); ",
	    dragElement,xto,yto);
	}
	public void TableSearch() throws InterruptedException, IOException
	{

   // WebElement Webtable=D8.findElement(By.xpath(".//*[@id='mainView']/div")); // Replace TableID with Actual Table ID or Xpath
	  WebElement Webtable=D8.findElement(By.xpath("//*[@id='table1']/tbody[1]"));
	  
    List<WebElement> TotalRowCount=Webtable.findElements(By.tagName("tr"));
    System.out.println("No. of Rows in the WebTable: "+TotalRowCount.size() +"Web Table:"+Webtable.isDisplayed());
    int rowcount=TotalRowCount.size()-3;
    // Now we will Iterate the Table and print the Values   
  // int RowIndex=1;
   // outerloop:
    //for(WebElement rowElement:TotalRowCount)
       
    for (int RowIndex=1;RowIndex<=rowcount;RowIndex++)
    {
      
    System.out.println("Traversing on Row and row count is  "  +RowIndex  +rowcount);
      //System.out.println(rowElement.getText());
   // List<WebElement> TotalColumnCount=rowElement.findElements(By.tagName("td"));
   // System.out.println("No. of columns in the WebTable:"+TotalColumnCount.size());
    
   // int ColumnIndex=1;
    
   // for(WebElement colElement:TotalColumnCount)
   // {
    //System.out.println("Row "+RowIndex+" Column "+ColumnIndex+" Data "+colElement.getText());
    //TestingCompany
   
    elem=D8.findElement(By.xpath(".//*[@id='table1']/tbody["+RowIndex+"]/tr[1]/td[4]"));
    String ServiceSubscriptionprocess=elem.getText();
    
    System.out.println("The element data is  "  + ServiceSubscriptionprocess + " Value of object is" +dCellData);
    
    elem=D8.findElement(By.xpath(".//*[@id='table1']/tbody["+RowIndex+"]/tr[1]/td[2]"));
    
    String ServiceName=elem.getText();
    System.out.println("Service Provider to be subscribed" +  ObjectSetVal);
    
   // if (ServiceSubscriptionprocess.equalsIgnoreCase(dCellData) && ServiceName.equalsIgnoreCase(ObjectSetVal))
    if (ServiceSubscriptionprocess.equalsIgnoreCase(dCellData))
    {
      System.out.println("Matched Service Name & process");
      try {
        ((JavascriptExecutor) D8).executeScript(
                "arguments[0].scrollIntoView(true);", elem);
        
    } catch (Exception e) {
      
      System.out.println("Exception in scrolling");
      Update_Report("failed");
    }
      elem.click();
      WebElement subscribeUnSubscibeButton=D8.findElement(By.xpath(".//*[@id='mainView']/div/div/div[1]/div[3]/button"));
      if(subscribeUnSubscibeButton.isEnabled())
      {
        Thread.sleep(8000);
        subscribeUnSubscibeButton.click();
      }
      System.out.println("Selected the Desired process");
    //  break outerloop;
    //}
    
    
  //  ColumnIndex=ColumnIndex+1;
    }
 //   RowIndex=RowIndex+1;
    }     
	}
	//Calender Method
public void selectJQueryDate(String date) {
  
  boolean breakIt = true;
  while (true) {
  breakIt = true;
  try {
     WebElement table = D8.findElement(By.className("gwt-DateBox ng-pristine ng-valid ng-valid-required"));
D8.switchTo().activeElement();
System.out.println("Table displayed " +table.isDisplayed());

List<WebElement> tableRows = table.findElements(By.tagName("tr"));
  for (WebElement row : tableRows) {
List<WebElement> cells = row.findElements(By.tagName("td"));

for (WebElement cell : cells) {
  System.out.println("Text is " +cell.getText());
  if (cell.getText().equals(date)) {
    
    cell.click();
    D8.switchTo().defaultContent();
    
 //   D8.findElement(By.linkText(date)).click();
  }
}
}
  } catch (Exception e) {
      if (e.getMessage().contains("element is not attached")) {
          breakIt = false;
      }
  }
  if (breakIt) {
      break;
  }

}

   
  }
		
	
// Method for Highlighting Objects
	public static void elementHighlight(WebElement element) throws InterruptedException {
		for (int i = 0; i < 2; i++) {
			JavascriptExecutor js = (JavascriptExecutor) D8;
		/*	js.executeScript(
					"arguments[0].setAttribute('style', arguments[1]);",
					element, "color: red; border: 1px groove green;");
			js.executeScript(
					"arguments[0].setAttribute('style', arguments[1]);",
					element, "");
			Thread.sleep(1000);*/
			
			 js.executeScript("arguments[0].style.border='1px groove green'", element);
       Thread.sleep(1000);
       js.executeScript("arguments[0].style.border=''", element);
		}
	}
	
	public static boolean isInteger(String input) {
		try {
			Integer.parseInt(input);
			return true;
		} catch (Exception e) {
			return false;
		}
	}
}
