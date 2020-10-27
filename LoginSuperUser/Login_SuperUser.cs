using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace LoginSuperUser
{
    public class Program
    {
        static void Main(string[] args)
        {
            //Open Chrome Driver
            IWebDriver driver = new ChromeDriver(@"C:\Users\Anneline\source\repos\");

            //Open Excel Spreadsheet to read data
            var excelApp = new Excel.Application();
            excelApp.Visible = true;
            Excel.Workbook xlWorkbook = excelApp.Workbooks.Open(@"C:\DATA\UserLoginData.xlsx");
            Excel._Worksheet xlWorksheet = (Excel._Worksheet)xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            for (int i = 2; i <= rowCount; i++) //Rowcount
            {

                for (int j = 1; j <= colCount; j++) //Columncount
                {
                    //Open CaMS Dev Environment
                    var UrlLink = (Excel.Range)xlWorksheet.Cells[i, j];
                    string WebLink = UrlLink.Value2.ToString();
                    driver.Navigate().GoToUrl(WebLink);
                    driver.Manage().Window.Maximize();
                    driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
                    j = j + 1;
                    
                    //Enter the Username
                    var UserN = (Excel.Range)xlWorksheet.Cells[i, j];
                    string Username = UserN.Value2.ToString();
                    IWebElement UserName = driver.FindElement(By.Id("Input_Email"));
                    UserName.SendKeys(Username);
                    j = j + 1;

                    //Enter the Password
                    var PasSWord = (Excel.Range)xlWorksheet.Cells[i, j];
                    string PassW = PasSWord.Value2.ToString();
                    IWebElement PassWOrd = driver.FindElement(By.Id("Input_Password"));
                    PassWOrd.SendKeys(PassW);
                    j = j + 1;

                    //Click on the Login Button
                    IWebElement LoginButton1 = driver.FindElement(By.CssSelector("#account > div:nth-child(4) > div > button"));
                    LoginButton1.Click();

                    //Click on the Logout Button
                    Actions action1 = new Actions(driver);
                    action1.MoveToElement(driver.FindElement(By.XPath("//*[@id='navbarDropdownProfile']/img"))).Build().Perform();
                    IWebElement manageProfile = driver.FindElement(By.XPath("//*[@id='navbarDropdownProfile']/img"));
                    manageProfile.Click();
                    IWebElement LogInoutButton2 = driver.FindElement(By.XPath("/html/body/div/main/nav/div/div[2]/ul/li[4]/div/a[3]"));
                    LogInoutButton2.Click();

                    //Click on the Login Button
                    IWebElement LoginButton3 = driver.FindElement(By.XPath("/html/body/div/div/div/div[1]/div/div[2]/div[2]/div[1]/div[2]/a"));
                    LoginButton3.Click();

                }

            }
            
            //Cleanup and Quit Excel Spreadsheet
            GC.Collect();
            GC.WaitForPendingFinalizers();            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);            //quit and release
            excelApp.Quit();
            Marshal.ReleaseComObject(excelApp);

            //Quit Browser
            driver.Quit();
        }
    }
}
