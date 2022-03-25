using Beezlabs.RPAHive.Lib;
using Beezlabs.RPAHive.Lib.Models;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System.Threading;

namespace Beezlabs.RPA.Bots
{
    public class GstRegistrationBotSix : RPABotTemplate
    {
        protected override void BotLogic(BotExecutionModel botExecutionModel)
        {
            ChromeOptions options = new ChromeOptions();
            options.BinaryLocation = "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe";
            options.DebuggerAddress = "localhost:9876";
            ChromeDriver chrome = new ChromeDriver(GetBotPath(), options);
            LogMessage(this.GetType().FullName, "Bot execution started");
            chrome.FindElementByXPath("//*[@id='bno']").SendKeys("");
            chrome.FindElementByXPath("//*[@id='bp_flrnum']").SendKeys("");
            chrome.FindElementByXPath("//*[@id='bp_bdname']").SendKeys("");
            chrome.FindElementByXPath("//*[@id='st']").SendKeys("");
            chrome.FindElementByXPath("//*[@id='loc']").SendKeys("");
            chrome.FindElementByXPath("").SendKeys("");
            var district = chrome.FindElementByXPath("//*[@id='cnty']");
            var districtSelect = new SelectElement(district);
            districtSelect.SelectByText("");
            chrome.FindElementByXPath("//*[@id='pncd']").SendKeys("");
            chrome.FindElementByXPath("//*[@id='stj']").Click();
            chrome.FindElementByXPath("//*[@id='stj']/option[2]").Click();
            chrome.FindElementByXPath("//*[@id='comcd']/option[2]").Click();
            chrome.FindElementByXPath("//*[@id='divcd']/option[2]").Click();
            chrome.FindElementByXPath("//*[@id='rgcd']/option[2]").Click();
            chrome.FindElementByXPath("//*[@id='bp_email']").Click();
            chrome.FindElementByXPath("//*[@id='tlphnoStd']").Clear();
            chrome.FindElementByXPath("//*[@id='bp_mobile']").SendKeys("");
            var natureOfPremises = chrome.FindElementByXPath("//*[@id='bp_buss_poss']");
            var natureOfPremisesSelect = new SelectElement(natureOfPremises);
            natureOfPremisesSelect.SelectByText("");
            var documentUpload = chrome.FindElementByXPath("//*[@id='bp_buss_poss']");
            var documentUploadSelect = new SelectElement(documentUpload);
            documentUploadSelect.SelectByText("");
            IWebElement uploadFile =  chrome.FindElementByXPath("//*[@id='newRegForm']/fieldset/div[3]/div/div/div[2]/div/data-upload-model/input");
            Thread.Sleep(2000);
            uploadFile.SendKeys("");
        }
    }
}