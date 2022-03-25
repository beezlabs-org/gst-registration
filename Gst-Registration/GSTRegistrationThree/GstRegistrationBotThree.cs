using Beezlabs.RPAHive.Lib;
using Beezlabs.RPAHive.Lib.Models;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.IO;
using System.Linq;
using System.Security;
using System.Threading;

namespace Beezlabs.RPA.Bots
{
    public class Details
    {
        public string iamA { get; set; }
        public string state { get; set; }
        public string district { get; set; }
        public string name { get; set; }
        public string PAN { get; set; }
        public string email { get; set; }
        public string mobileNo { get; set; }
        public string tradeName { get; set; }
        public string constOfBusiness { get; set; }
        public string dateOfCommencement { get; set; }
        public string typeOfRegistration { get; set; }
        public string registrationNo { get; set; }
        public string dateOfRegistration { get; set; }
        public string CIN { get; set; }
    }
    public class GstRegistrationBotThree : RPABotTemplate
    {
        public Details details = new Details();
        JArray values;
        protected override void BotLogic(BotExecutionModel botExecutionModel)
        {
            try
            {
                if (botExecutionModel.proposedBotInputs["GstData"] != null)
                {
                    JObject sap = (JObject)botExecutionModel.proposedBotInputs["GstData"].value;
                    details = sap.ToObject<Details>();
                }
                //Sharepoint sharepoint = new Sharepoint();
                //string webUrl = botExecutionModel.proposedBotInputs["WebUrl"].value.ToString();
                //string docLibName = botExecutionModel.proposedBotInputs["DocLibname"].value.ToString();
                //string fileName = botExecutionModel.proposedBotInputs["FileName"].value.ToString();
                //string IdentityName = botExecutionModel.proposedBotInputs["Credentials"].value.ToString();
                //string userName = botExecutionModel.identityList.SingleOrDefault(cred => cred.name.Equals(IdentityName)).credential.basicAuth.username.ToString(); 
                //string password = botExecutionModel.identityList.SingleOrDefault(cred => cred.name.Equals(IdentityName)).credential.basicAuth.password.ToString();
                LogMessage(this.GetType().FullName, "Bot started");
                if (botExecutionModel.proposedBotInputs["DetailsOfPromoter"] != null)
                {
                    string val = botExecutionModel.proposedBotInputs["details"].value.ToString();
                    values = (JArray)JsonConvert.DeserializeObject(val);                   
                    LogMessage(this.GetType().FullName, "Deserialized values");
                }
                string secondOtp = botExecutionModel.proposedBotInputs["SecondMobileOTP"].value.ToString();
                LogMessage(this.GetType().FullName, "Fetched mobile otp");
                string trn = botExecutionModel.proposedBotInputs["TRN"].value.ToString();
                LogMessage(this.GetType().FullName, "Fetched TRN "+trn);
                ChromeOptions options = new ChromeOptions();
                options.BinaryLocation = "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe";
                options.DebuggerAddress = "localhost:9876";
                ChromeDriver chrome = new ChromeDriver(GetBotPath(), options);
                chrome.FindElementByXPath("//*[@id='mobile_otp']").SendKeys(secondOtp);
                chrome.FindElementByXPath("/html/body/div[2]/div[2]/div/div[2]/div/div[2]/div/form/div/div/button").Click();
                Thread.Sleep(3000);
                chrome.FindElementByXPath("/html/body/div[2]/div[1]/div/div[3]/div[2]/div/div/table/tbody/tr/td[6]/button").Click();
                Thread.Sleep(3000);
                IWebElement element = chrome.FindElementByXPath("//*[@id='tnm']");
                IJavaScriptExecutor js = (IJavaScriptExecutor)chrome;
                js.ExecuteScript("arguments[0].scrollIntoView();",element);
                chrome.FindElementByXPath("//*[@id='tnm']").SendKeys(details.tradeName);
                LogMessage(this.GetType().FullName, "Entered trade name" + details.tradeName);
                Thread.Sleep(3000);
                var constOfBusiness = chrome.FindElementByXPath("//*[@id='bd_ConstBuss']");
                var constOfBusinessSelect = new SelectElement(constOfBusiness);
                LogMessage(this.GetType().FullName, "howered on constOfBusinessSelect");
                constOfBusinessSelect.SelectByText(details.constOfBusiness); //Proprietorship //"Private Limited Company"
                var registrationReason = chrome.FindElementByXPath("//*[@id='bd_rsl']");
                var registrationReasonSelect = new SelectElement(registrationReason);
                registrationReasonSelect.SelectByText("Voluntary Basis");
                string date = DateTime.Today.ToString("dd/MM/yyyy");
                Thread.Sleep(3000);
                chrome.FindElementByXPath("//*[@id='bd_cmbz']").Click();
                chrome.FindElementByXPath("//*[@id='bd_cmbz']").SendKeys(date);
                LogMessage(this.GetType().FullName, "date of commencement has been entered " +date);
                chrome.FindElementByXPath("//*[@id='lib']").Click();
                chrome.FindElementByXPath("//*[@id='lib']").SendKeys(date);
                LogMessage(this.GetType().FullName, "date of commencement has been entered " + date);
                if(details.constOfBusiness.Equals("Private Limited Company"))
                {
                    var typeOfRegistration = chrome.FindElementByXPath("//*[@id='exty']");
                    var typeOfRegistrationSelect = new SelectElement(typeOfRegistration);
                    typeOfRegistrationSelect.SelectByText("Corporate Identity Number / Foreign Company Registration Number");//Limited Liability Partnership / Foreign Limited Liability Partnership Identification Number
                    chrome.FindElementByXPath("//*[@id='exno']").SendKeys(details.CIN);
                    chrome.FindElementByXPath("//*[@id='exdt']").SendKeys(details.dateOfCommencement);
                    Thread.Sleep(2000);
                    chrome.FindElementByXPath("//*[@id='newRegForm']/fieldset/div/div[7]/div/div[4]/button[1]").Click(); // add
                    Thread.Sleep(2000);
                    var proof = chrome.FindElementByXPath("//*[@id='bd_up_type']");
                    js.ExecuteScript("arguments[0].scrollIntoView();", proof);
                    var proofSelect = new SelectElement(proof);
                    proofSelect.SelectByText("Certificate of Incorporation");
                    SecureString secureString = new SecureString();
                    string workingDirectory = GetWorkingDirectory().ToString();
                    //foreach (char c in password.ToCharArray())
                    //{
                    //    secureString.AppendChar(c);
                    //}
                    //System.Net.ICredentials SPCred = new SharePointOnlineCredentials(userName, secureString);
                    //sharepoint.DownloadFileViaRestAPI(webUrl, SPCred, docLibName, fileName, workingDirectory);
                    IWebElement uploadProof = chrome.FindElementByXPath("/html/body/div[2]/div/div/div[3]/form/fieldset/div[2]/div/div/div[2]/fieldset/data-file-model/input");
                    //LogMessage(this.GetType().FullName, "upload proof element selected");
                    //Thread.Sleep(4000);
                    //string path = Path.Combine(workingDirectory, fileName);
                    //LogMessage(this.GetType().FullName, "working directory and filename " + workingDirectory + fileName);
                    string path = null;
                    uploadProof.SendKeys(path);
                    Thread.Sleep(2000);
                }                
                chrome.FindElementByXPath("//*[@id='newRegForm']/div/div/button[2]").Click(); //save and continue
                Success("Bot executed successfully");
            }
            catch (Exception e)
            {
                Failure("Bot execution failed " + e.Message);
            }
        }
    }
}