using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using ClosedXML.Excel;
using System.Threading;
using System.IO;

namespace RPAChallengeSelenium
{
    public class Program
    {
        public static void DeletaArquivoSeExistente(string file)
        {
            if (File.Exists(file))
            {
                File.Delete(file);
            }
        }

        static void Main(string[] args)
        {
            DeletaArquivoSeExistente(@"C:\Users\fiska\Downloads\challenge.xlsx");

            var driver = new ChromeDriver();

            driver.Navigate().GoToUrl("https://www.rpachallenge.com/");

            driver.Manage().Window.Maximize();

            driver.FindElement(By.XPath("/html/body/app-root/div[2]/app-rpa1/div/div[1]/div[6]/a")).Click();
            Thread.Sleep(5000);
            driver.FindElement(By.XPath("/html/body/app-root/div[2]/app-rpa1/div/div[1]/div[6]/button")).Click();

            var xls = new XLWorkbook(@"C:\Users\fiska\Downloads\challenge.xlsx");
 
            var planilha = xls.Worksheets.First(w => w.Name == "Sheet1");
            var totalLinhas = planilha.Rows().Count();


            // primeira linha é o cabecalho
            for (int l = 2; l <= 11; l++)
            {

                var firstName = planilha.Cell($"A{l}").Value.ToString();
                var lastName = planilha.Cell($"B{l}").Value.ToString();
                var companyName = planilha.Cell($"C{l}").Value.ToString();
                var roleInCompany = planilha.Cell($"D{l}").Value.ToString();
                var address = planilha.Cell($"E{l}").Value.ToString();
                var email = planilha.Cell($"F{l}").Value.ToString();
                var phoneNumber = planilha.Cell($"G{l}").Value.ToString();

                driver.FindElement(By.XPath("//input[@ng-reflect-name='labelPhone']")).SendKeys(phoneNumber);
                driver.FindElement(By.XPath("//input[@ng-reflect-name='labelEmail']")).SendKeys(email);
                driver.FindElement(By.XPath("//input[@ng-reflect-name='labelFirstName']")).SendKeys(firstName);
                driver.FindElement(By.XPath("//input[@ng-reflect-name='labelLastName']")).SendKeys(lastName);
                driver.FindElement(By.XPath("//input[@ng-reflect-name='labelAddress']")).SendKeys(address);
                driver.FindElement(By.XPath("//input[@ng-reflect-name='labelRole']")).SendKeys(roleInCompany);
                driver.FindElement(By.XPath("//input[@ng-reflect-name='labelCompanyName']")).SendKeys(companyName);

                driver.FindElement(By.XPath("//input[@class='btn uiColorButton']")).Click();
            }

            var message = driver.FindElement(By.XPath("/html/body/app-root/div[2]/app-rpa1/div/div[2]/div[2]")).Text.ToString();
            Console.WriteLine(message);
        }
    }
}
