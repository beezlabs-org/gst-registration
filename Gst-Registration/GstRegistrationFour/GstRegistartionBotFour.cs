using Beezlabs.RPAHive.Lib;
using Beezlabs.RPAHive.Lib.Models;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security;

namespace Beezlabs.RPA.Bots
{
    public class DetailsOfPromoter
    {
        public string firstName { get; set; }
        public string middleName { get; set; }
        public string lastName { get; set; }
        public string fatherFirstName { get; set; }
        public string fatherMiddleName { get; set; }
        public string fatherLastName { get; set; }
        public string dob { get; set; }
        public string mobileNumber { get; set; }
        public string email { get; set; }
        public string gender { get; set; }
        public string designation { get; set; }
        public string DIN { get; set; }
        public string PAN { get; set; }
        public string flatNo { get; set; }
        public string floorNo { get; set; }
        public string nameofBuilding { get; set; }
        public string Street { get; set; }
        public string city { get; set; }
        public string state { get; set; }
        public string district { get; set; }
        public string zipCode { get; set; }
        public string documentUploadName { get; set; }
    }
    public class GstRegistartionBotFour : RPABotTemplate
    {
        public DetailsOfPromoter detailsOfPromoter = new DetailsOfPromoter();
        JArray values;
        protected override void BotLogic(BotExecutionModel botExecutionModel)
        {
            try
            {
                System.Diagnostics.Debugger.Launch();
                System.Diagnostics.Debugger.Break();
                Sharepoint sharepoint = new Sharepoint();
                string webUrl = botExecutionModel.proposedBotInputs["WebUrl"].value.ToString();
                string docLibName = botExecutionModel.proposedBotInputs["DocLibname"].value.ToString();
                string fileName = botExecutionModel.proposedBotInputs["FileName"].value.ToString();
                string IdentityName = botExecutionModel.proposedBotInputs["Credentials"].value.ToString();
                string userName = botExecutionModel.identityList.SingleOrDefault(cred => cred.name.Equals(IdentityName)).credential.basicAuth.username.ToString();
                string password = botExecutionModel.identityList.SingleOrDefault(cred => cred.name.Equals(IdentityName)).credential.basicAuth.password.ToString();
                ChromeOptions options = new ChromeOptions();
                options.BinaryLocation = "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe";
                options.DebuggerAddress = "localhost:9876";
                ChromeDriver chrome = new ChromeDriver(GetBotPath(), options);
                LogMessage(this.GetType().FullName, "Bot execution started");
                int v = 0;
                if (botExecutionModel.proposedBotInputs["DetailsOfPromoter"] != null)
                {
                    string val = botExecutionModel.proposedBotInputs["DetailsOfPromoter"].value.ToString();
                    values = (JArray)JsonConvert.DeserializeObject(val);
                    v = values.Count;
                    LogMessage(this.GetType().FullName, "values count " + v);
                }
                for (int i = 0; i < v; i++)
                {
                    if (values[i]["firstName"].ToString() != null)
                    {
                        chrome.FindElementByXPath("//*[@id='fnm']").SendKeys(values[i]["firstName"].ToString());
                    }
                    else
                    {

                    }
                    if (values[i]["middleName"].ToString() != null)
                    {
                        chrome.FindElementByXPath("//*[@id='pd_mname']").SendKeys(values[i]["middleName"].ToString());
                    }
                    else
                    {

                    }
                    if (values[i]["lastName"].ToString() != null)
                    {
                        chrome.FindElementByXPath("//*[@id='pd_lname']").SendKeys(values[i]["lastName"].ToString());
                    }
                    else
                    {

                    }
                    if (values[i]["fatherFirstName"].ToString() != null)
                    {
                        chrome.FindElementByXPath("//*[@id='ffname']").SendKeys(values[i]["fatherFirstName"].ToString());
                    }
                    else
                    {

                    }
                    if (values[i]["fatherMiddleName"].ToString() != null)
                    {
                        chrome.FindElementByXPath("//*[@id='pd_fmname']").SendKeys(values[i]["fatherMiddleName"].ToString());
                    }
                    else
                    {

                    }
                    if (values[i]["fatherLastName"].ToString() != null)
                    {
                        chrome.FindElementByXPath("//*[@id='pd_flname']").SendKeys(values[i]["fatherLastName"].ToString());
                    }
                    else
                    {

                    }
                    if (values[i]["dob"].ToString() != null)
                    {
                        chrome.FindElementByXPath("//*[@id='dob']").SendKeys(values[i]["dob"].ToString());
                    }
                    else
                    {

                    }
                    if (values[i]["gender"].ToString().ToLower().Equals("Male"))
                    {
                        chrome.FindElementByXPath("//*[@id='radiomale']").Click();
                    }
                    else if (values[i]["gender"].ToString().ToLower().Equals("Female"))
                    {
                        chrome.FindElementByXPath("//*[@id='radiofemale']").Click();
                    }
                    else
                    {
                        chrome.FindElementByXPath("//*[@id='radiotrans']").Click();
                    }
                    if (values[i]["designation"].ToString() != null)
                    {
                        chrome.FindElementByXPath("//*[@id='dg']").SendKeys(values[i]["designation"].ToString());
                    }
                    else
                    {

                    }
                    if (values[i]["DIN"].ToString() != null)
                    {
                        chrome.FindElementByXPath("//*[@id='din']").SendKeys(values[i]["DIN"].ToString());
                    }
                    else
                    {

                    }
                    if (values[i]["flatNo"].ToString() != null)
                    {
                        chrome.FindElementByXPath("//*[@id='pd_bdnum']").SendKeys(values[i]["flatNo"].ToString());
                    }
                    else
                    {

                    }
                    if (values[i]["nameofBuilding"].ToString() != null)
                    {
                        chrome.FindElementByXPath("//*[@id='pd_bdname']").SendKeys(values[i]["nameofBuilding"].ToString());
                    }
                    else
                    {

                    }
                    if (values[i]["Street"].ToString() != null)
                    {
                        chrome.FindElementByXPath("//*[@id='pd_road']").SendKeys(values[i]["Street"].ToString());
                    }
                    else
                    {

                    }
                    if (values[i]["city"].ToString() != null)
                    {
                        chrome.FindElementByXPath("//*[@id='pd_locality']").SendKeys(values[i]["city"].ToString());
                    }
                    else
                    {

                    }
                    var country = chrome.FindElementByXPath("//*[@id='pd_cntry']");
                    var countrySelect = new SelectElement(country);
                    countrySelect.SelectByText("India");
                    if (values[i]["district"].ToString() != null)
                    {
                        var district = chrome.FindElementByXPath("//*[@id='dst']");
                        var districtSelect = new SelectElement(district);
                        LogMessage(this.GetType().FullName, "howered on district");
                        districtSelect.SelectByText(detailsOfPromoter.district);
                    }
                    else
                    {

                    }
                    SecureString secureString = new SecureString();
                    string workingDirectory = GetWorkingDirectory().ToString();
                    foreach (char c in password.ToCharArray())
                    {
                        secureString.AppendChar(c);
                    }
                    System.Net.ICredentials SPCred = new SharePointOnlineCredentials(userName, secureString);
                    sharepoint.DownloadFileViaRestAPI(webUrl, SPCred, docLibName, values[i]["documentUploadName"].ToString(), workingDirectory);
                    chrome.FindElementByXPath("//*[@id='pncd']").SendKeys(values[i]["zipCode"].ToString());
                    IWebElement docUpload = chrome.FindElementByXPath("//*[@id='newRegForm']/div[2]/fieldset/div[4]/div/div/div[1]/data-file-model/input")
                    string path = Path.Combine(workingDirectory, values[i]["documentUploadName"].ToString());
                    docUpload.SendKeys(path);
                   // chrome.FindElementByXPath("//*[@id='pri_auth']").Click(); //toggle, also authorized signatory
                    chrome.FindElementByXPath("//*[@id='newRegForm']/div[2]/div[2]/div/button[2]").Click();//Add new click
                }
                chrome.FindElementByXPath("//button[contains(text(),'Save & Continue')]").Click(); //save and continue
                Success("Bot Executed succesfully");
            }
            catch (Exception ex)
            {
                Failure($"Bot Executed failed due to {ex.Message}");
                throw;
            }
        }

    }
}