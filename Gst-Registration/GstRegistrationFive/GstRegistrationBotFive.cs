using Beezlabs.RPAHive.Lib;
using Beezlabs.RPAHive.Lib.Models;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System.Threading;

namespace GstRegistrationFive
{
    public class GstRegistrationBotFive : RPABotTemplate
    {
        protected override void BotLogic(BotExecutionModel botExecutionModel)
        {
            ChromeOptions options = new ChromeOptions();
            options.BinaryLocation = "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe";
            options.DebuggerAddress = "localhost:9876";
            ChromeDriver chrome = new ChromeDriver(GetBotPath(), options);
            LogMessage(this.GetType().FullName, "Bot execution started");
            //forloop
            chrome.FindElementByXPath("//*[@id='fnm']").SendKeys("");
            chrome.FindElementByXPath("//*[@id='as_mname']").SendKeys("");
            chrome.FindElementByXPath("//*[@id='as_lname']").SendKeys("");
            chrome.FindElementByXPath("//*[@id='ffname']").SendKeys("");
            chrome.FindElementByXPath("//*[@id='as_fmname']").SendKeys("");
            chrome.FindElementByXPath("//*[@id='as_flname']").SendKeys("");
            chrome.FindElementByXPath("//*[@id='dob']").Click();
            chrome.FindElementByXPath("//*[@id='dob']").SendKeys("");
            chrome.FindElementByXPath("//*[@id='mbno']").SendKeys("");
            chrome.FindElementByXPath("//*[@id='em']").SendKeys("");
            chrome.FindElementByXPath("//*[@id='radiomale']").Click();
            chrome.FindElementByXPath("//*[@id='dg']").SendKeys("");
            chrome.FindElementByXPath("//*[@id='pan']").SendKeys("");
            chrome.FindElementByXPath("//*[@id='bno']").SendKeys("");
            chrome.FindElementByXPath("//*[@id='as_flrnum']").SendKeys("");
            chrome.FindElementByXPath("//*[@id='as_bdname']").SendKeys("");
            chrome.FindElementByXPath("//*[@id='st']").SendKeys("");
            chrome.FindElementByXPath("//*[@id='loc']").SendKeys("");
            var country = chrome.FindElementByXPath("//*[@id='cnty']");
            var countrySelect = new SelectElement(country);
            countrySelect.SelectByText("India");
            var state = chrome.FindElementByXPath("//*[@id='stcd']");
            var stateSelect = new SelectElement(state);
            LogMessage(this.GetType().FullName, "howered on district");
            stateSelect.SelectByText("");
            var district = chrome.FindElementByXPath("//*[@id='stcd']");
            var districtSelect = new SelectElement(district);
            LogMessage(this.GetType().FullName, "howered on district");
            districtSelect.SelectByText("");
            chrome.FindElementByXPath("//*[@id='pncd']").SendKeys("");
            var authorizedSignatory = chrome.FindElementByXPath("//*[@id='stcd']");
            var authorizedSignatorySelect = new SelectElement(authorizedSignatory);
            LogMessage(this.GetType().FullName, "howered on district");
            authorizedSignatorySelect.SelectByText("");
            chrome.FindElementByXPath("//*[@id='newRegForm']/div[2]/fieldset[2]/div/div[2]/div/fieldset/div/data-file-model/input").SendKeys("");
            Thread.Sleep(2000);
            chrome.FindElementByXPath("//*[@id='newRegForm']/div[2]/fieldset[2]/div/div[2]/div/div[1]/data-file-model/input").SendKeys("");
            chrome.FindElementByXPath("//*[@id='newRegForm']/div[2]/div[2]/div/button[2]").Click(); //Add new
            chrome.FindElementByXPath("//*[@id='newRegForm']/div[2]/div[2]/div/button[3]").Click();// save and continue
            Thread.Sleep(3000);
            chrome.FindElementByXPath("//*[@id='newRegForm']/div/div/button[2]").Click();
        }
    }
}