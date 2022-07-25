using System;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using System.Threading;

namespace control
{
    public class AuthBrowser
    {
        public void Browser()
        {
            Fields.login = "LOGIN";
            Fields.password = "PASSWORD";

            var driverService = ChromeDriverService.CreateDefaultService();                                         
            driverService.HideCommandPromptWindow = true;                                                           
            var driver = new ChromeDriver(driverService, new ChromeOptions());                                     
            ChromeOptions options = new ChromeOptions();                                                            
            driver.Manage().Window.Maximize();                                                                                            
            options.AddArguments("--incognito");                                                                              

            driver.Navigate().GoToUrl("https://web.suite.com/#/login/");
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(3);                                      

            IWebElement logon = driver.FindElement(By.XPath("//*[@id='input_0']"));                                 
            logon.SendKeys(Fields.login);                                                                           

            IWebElement pass = driver.FindElement(By.XPath("//*[@id='input_1']"));                                  
            pass.SendKeys(Fields.password);                                                                         

            IWebElement BTNlogin = driver.FindElement(By.CssSelector("#login-btn-login"));                          
            BTNlogin.Click();                                                                                       
            Thread.Sleep(1000);

            IWebElement CreateWaybill = driver.FindElement(By.CssSelector("#sidebar-btn-invoice-create > span"));   
            CreateWaybill.Click();

            IWebElement Descript = driver.FindElement(By.CssSelector("#mat-input-12"));                                                                  
            Descript.SendKeys(Fields.Description);

            IWebElement Cargo = driver.FindElement(By.CssSelector("#mat-input-13"));                                                       
            Cargo.Clear();
            Cargo.SendKeys(Fields.Cost);

            IWebElement Weight = driver.FindElement(By.CssSelector("#mat-input-8"));                                                           
            Weight.SendKeys(Fields.TotalWeight);

            IWebElement Length = driver.FindElement(By.CssSelector("#mat-input-9"));                                                        
            Length.SendKeys(Fields.ItemLength);

            IWebElement Width = driver.FindElement(By.CssSelector("#mat-input-10"));                                                        
            Width.SendKeys(Fields.ItemWidth);

            IWebElement Height = driver.FindElement(By.CssSelector("#mat-input-11"));                                                       
            Height.SendKeys(Fields.ItemHeight);
            Thread.Sleep(500);

            driver.ExecuteScript("window.scrollTo(0, 250)");

            IWebElement Sender = driver.FindElement(By.CssSelector("#mat-radio-9 > label > div.mat-radio-label-content"));                  
            Sender.Click();

            IWebElement Cashless = driver.FindElement(By.CssSelector("#mat-radio-12 > label > div.mat-radio-label-content"));               
            Cashless.Click();

            IWebElement Info = driver.FindElement(By.CssSelector("#mat-input-4"));                                                         
            Info.SendKeys(Fields.Description);

            driver.ExecuteScript("window.scrollTo(0, 250)");                                                                                
            Thread.Sleep(500);

            IWebElement Next = driver.FindElement(By.CssSelector("#edit-invoice-btn-create"));                                             
            Next.Click();

            driver.ExecuteScript("window.scrollTo(0, 250)");

            IWebElement Company = driver.FindElement(By.CssSelector("#mat-input-5"));                                                      
            Company.SendKeys(Fields.Recipient);
            Company.Click();
            Thread.Sleep(500);

            IWebElement selectRecipient = driver.FindElement(By.ClassName("mat-option-text"));                              
            selectRecipient.Click();

            driver.ExecuteScript("window.scrollTo(0, 250)");                                                                
            Thread.Sleep(500);

            IWebElement RecipientLocality = driver.FindElement(By.Id("mat-input-29"));                                             
            RecipientLocality.SendKeys(Fields.Locality);
            RecipientLocality.Click();
            Thread.Sleep(3000);

            IWebElement selectRecipientLocality = driver.FindElement(By.ClassName("mat-option-text"));                      
            selectRecipientLocality.Click();

            IWebElement RecipientBranch = driver.FindElement(By.Id("mat-input-30"));                                                
            RecipientBranch.SendKeys(Fields.Branch);
            RecipientBranch.Click();
            Thread.Sleep(500);

            IWebElement selectRecipientBranch = driver.FindElement(By.ClassName("mat-option-text"));                        
            selectRecipientBranch.Click();
            Thread.Sleep(3000);

            driver.Close();

            Fields.row++;
            
            FileSelected selected = new FileSelected();
            selected.xlsxSelected();            
        }
    }
}
