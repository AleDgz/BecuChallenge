using NUnit.Framework;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Threading;
using System.Timers;
using System.IO;
using OfficeOpenXml;


namespace NUnit.Tests1
{
    [TestFixture]
    public class TestClass
    {
        public IWebDriver driver;
        

        [OneTimeSetUp]
         public void Open()
         {
             //creating an objet from chromedriver
             driver = new ChromeDriver();
             driver.Manage().Window.Maximize();
             driver.Url = "https://www.BECU.org";
             driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(20);
        }

        [Test]
        public void TestMethod()
        {
            //find "Loan and Mortage" link and click
            driver.FindElement(By.XPath("//a[text() ='Loans & Mortgages']")).Click();

            //find "Auto Loans" and click
            driver.FindElement(By.XPath("//li//a[@title = 'Auto Loans']")).Click();

            //find calculator
            driver.FindElement(By.XPath("//button[contains(@data-bs-target,'How-much-ve')]")).SendKeys(Keys.Enter);

            //enter data
            driver.SwitchTo().Frame(driver.FindElement(By.XPath("//iframe[@title = 'Auto_What vehicle can I afford']")));
            Thread.Sleep(3000);
            
        }

        [Test]
        public void TestMethod2()
        {
          IWebElement payment = driver.FindElement(By.XPath("//input[@name = 'Auto_MonthlyPayment']"));
          IWebElement dPayment= driver.FindElement(By.XPath("//input[@name='Global_AutoDownPayment']"));
          IWebElement lTerm = driver.FindElement(By.XPath("//input[@name = 'Global_AutoLoanTerm']"));
          IWebElement iRate = driver.FindElement(By.XPath("//input[@name = 'Global_AutoInterestRate']"));

            
            using (ExcelPackage paquete = new ExcelPackage())
            {
                //load the excel document
                using (FileStream flujo = File.OpenRead(@"C:\Users\ale-d\source\repos\NUnit.Tests1\Libro1.xlsx"))
                { paquete.Load(flujo);}
                // First page of the document
                ExcelWorksheet hoja1 = paquete.Workbook.Worksheets.First();

                for (int i = 2; i <= 5; i++)
                {
                    String text = driver.FindElement(By.XPath("//span[@id = 'lf_answer']//span")).GetAttribute("innerText");
                    int numFila = i;

                    // Obtain value of cells an colls
                    string Month = hoja1.Cells[numFila, 2].Text;
                    string Down = hoja1.Cells[numFila, 3].Text;
                    string Loan = hoja1.Cells[numFila, 4].Text;
                    string Interest = hoja1.Cells[numFila, 5].Text;

                    //1st test values
                    payment.SendKeys(Keys.Control + "a");
                    payment.SendKeys(Keys.Delete);
                    payment.SendKeys(Month);

                    Thread.Sleep(2000);
                    dPayment.SendKeys(Keys.Control + "a");
                    dPayment.SendKeys(Keys.Delete);
                    dPayment.SendKeys(Down);

                    Thread.Sleep(2000);
                    lTerm.SendKeys(Keys.Control + "a");
                    lTerm.SendKeys(Keys.Delete);
                    lTerm.SendKeys(Loan);

                    Thread.Sleep(2000);
                    iRate.SendKeys(Keys.Control + "a");
                    iRate.SendKeys(Keys.Delete);
                    iRate.SendKeys(Interest);
                    iRate.SendKeys(Keys.Enter);

                    Thread.Sleep(3000);
                    String text2 = driver.FindElement(By.XPath("//span[@id = 'lf_answer']//span")).GetAttribute("innerText");
                    if (text.Equals(text2))
                    {
                        Console.WriteLine("You have the same loan = " + text2);
                    }
                    else {
                        Console.WriteLine("You have differente loan = " + text2);
                    }

                }
            }
        }
       
    }
}