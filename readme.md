Demo_SeleniumWebDriverFramework 🚀

A modular and data-driven automation framework built using Java + Selenium WebDriver. This project supports cross-browser testing, integrates with TestNG, JUnit, and Jenkins, and leverages Excel-based keyword-driven data for flexible test execution. Ideal for scalable UI testing and CI/CD pipelines.

🧰 Features

🔹 Keyword-driven and data-driven architecture

🔹 Browser compatibility: Chrome, Firefox, Safari, IE, PhantomJS

🔹 Integration with popular tools: TestNG, JUnit, Jenkins

🔹 Excel-based object repository and test data

🔹 Extensible structure for modular test cases

🔹 Supports CI pipelines for automated execution

🧪 Tech Stack

Java

Selenium WebDriver

Apache POI (Excel)

TestNG / JUnit

Jenkins

![Java](https://img.shields.io/badge/Java-ED8B00?style=for-the-badge&logo=java&logoColor=white)
![Selenium](https://img.shields.io/badge/Selenium-43B02A?style=for-the-badge&logo=selenium&logoColor=white)
![TestNG](https://img.shields.io/badge/TestNG-EF4035?style=for-the-badge)
![Jenkins](https://img.shields.io/badge/Jenkins-D24939?style=for-the-badge&logo=jenkins&logoColor=white)


The Framework uses open source Selenium Webdriver for executing automated scripts.
    URL for Reference : http://www.seleniumhq.org/projects/webdriver/
    
    
    2 FRAMEWORK SETUP:
    
    Please ensure that the file path are not hyperlink under "Location" column in Config_Framework.xls file. If 'yes' then just remove hyperlink for all file path which are mentioned in config_framework.xls file.

 
For Browser Type Framework supports  Test Execution on IE & Firefox, Google chrome and Safari

FF : Firefox
IE : Internet Explorer
GC : Google Chrome.
SA : Safari
PJS : Phantom Js  Headless browser

Note : For IE setting : 
1.	On IE 7 or higher on Windows Vista or Windows 7, you must set the Protected Mode settings for each zone to be the same value. The value can be on or off, as long as it is the same for every zone. To set the Protected Mode settings, choose “Internet Options…” from the Tools menu, and click on the Security tab. For each zone, there will be a check box at the bottom of the tab labeled “Enable Protected Mode”.

2.	"set IE :  Tools >> Internet options >> Advanced >> Security - 'Allow active content to run in files on my computer' ,this option must be selected.

3 For chrome download the latest chrome driver (https://sites.google.com/a/chromium.org/chromedriver/downloads)

4. For Firefox download firefox gecko driver


b. Object Repository Excel:

Collect the properties of objects and define it in
Object_Repository.xls like below.
The various locator types are explained below with examples for each. 
1.	By Id (e.g id=j_username)
2.	By xpath (e.g. Xpath=//* or tag name [@id/class/text]/) (absolute or relative xpath)
3.	By link Text (e.g link=Continue)
4.	By Name (e.g name=username )
5.	By CSS(e.g css=input[name="username"] )



The object repository for firefox can be added using firebug plugin for firebug usage refer url https://getfirebug.com/

Also Install firepath after firebug to get Xpath refer url
https://addons.mozilla.org/en-US/firefox/addon/firepath/

The Details all the keywords and Syntax are in seperate document.
Selenium_KDF_Keywords.xls
While importing test data give the relative path (i.e \\TestData\\{Foldername}\\TestData.xls or \\TestData\\TestData.xls


Step 1: Copy the "Selenium franework" folder having all excels on your machine it can be any drive C: Drive or D: 

Step 2: Update the config file  according to your path.

 
Step 3: Update the test data file present according to your test scripts path.

TEST EXECUTION:

"RUN" The Testcases by clicking as junit or Testng from any java ide ( eclipse, VS code )

This Framework can be also intergrated with Jenkins CI and run in headless mode.


For setup instructions and usage details, please refer to the 👉 SeleniumFramework_Quick_Start_Guide_V1.doc
