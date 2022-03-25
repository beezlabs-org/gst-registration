using Beezlabs.RPAHive.Lib;
using Beezlabs.RPAHive.Lib.Models;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading;
using System.Windows.Forms;

namespace Beezlabs.RPA.Bots
{
    class HeaderValue
    {
        public string key { get; set; } = null;
        public string type { get; set; } = null;
        public string items { get; set; } = null;
        public string label { get; set; } = null;
        public int length { get; set; } = 0;
        public string outline { get; set; } = null;
        public string cssClass { get; set; } = null;
        public bool required { get; set; } = false;
        public string fileProps { get; set; } = null;
        public string description { get; set; } = null;
        public string selectProps { get; set; } = null;
        public string defaultValue { get; set; } = null;
        public string flowVariable { get; set; } = null;
        public string deepStructure { get; set; } = null;

        public HeaderValue(string key, string type, string label, int length, bool required, string description, string defaultValue)
        {
            this.key = key;
            this.type = type;
            this.label = label;
            cssClass = "xs12";
            this.required = required;
            this.description = description;
            this.defaultValue = defaultValue;
        }

    }
    public class GstRegistrationBotOne : RPABotTemplate
    {
        JArray values;
        protected override void BotLogic(BotExecutionModel botExecutionModel)
        {
            try
            {
                System.Diagnostics.Debugger.Launch();
                System.Diagnostics.Debugger.Break();
                System.Diagnostics.Process process = new System.Diagnostics.Process();
                string userdataPath = GetBotPath() + "UserData\\";
                System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo();
                startInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
                startInfo.FileName = "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe";
                string GstUrl = "https://reg.gst.gov.in/registration/";
                startInfo.Arguments = $"{GstUrl} --new-window -remote-debugging-port=9876 --user-data-dir=\"C:\\Selenium\\helloworld\"";
                process.StartInfo = startInfo;
                process.Start();
                ChromeOptions options = new ChromeOptions();
                options.BinaryLocation = "C:\\Program Files(x86)\\Google\\Chrome\\Application\\chrome.exe";
                options.DebuggerAddress = "localhost:9876";
                ChromeDriver chrome = new ChromeDriver(GetBotPath(), options);
                chrome.Manage().Window.Maximize();
                Thread.Sleep(3000);
                //if (botExecutionModel.proposedBotInputs["details"] != null)
                //{
                //    string val = botExecutionModel.proposedBotInputs["details"].value.ToString();
                //    values = (JArray)JsonConvert.DeserializeObject(val);
                //}
                // WebDriverWait wait = new WebDriverWait(chrome, TimeSpan.FromSeconds(300));
                //IWebElement FindElement = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.Name("//*[@id='applnType']")));
                var appicationType = chrome.FindElementByXPath("//*[@id='applnType']");
                LogMessage(this.GetType().FullName, "application type found");
                var appicationTypeSelect = new SelectElement(appicationType);
                LogMessage(this.GetType().FullName, "howered on appln type");
                appicationTypeSelect.SelectByText("Taxpayer");
                LogMessage(this.GetType().FullName, "selected taxpayer");
                var state = chrome.FindElementByXPath("//*[@id='applnState']");
                var statevalue = new SelectElement(state);
                statevalue.SelectByText("Tamil Nadu");
                Thread.Sleep(3000);
                var district = chrome.FindElementByXPath("//*[@id='applnDistr']");
                var districtValue = new SelectElement(district);
                districtValue.SelectByText("Chennai");
                chrome.FindElementByXPath("//*[@id='bnm']").SendKeys("ADDANKI DAVIDRAJ STEVE RICHARD");
                chrome.FindElementByXPath("//*[@id='pan_card']").SendKeys("GHZPS7414C");
                chrome.FindElementByXPath("//*[@id='email']").SendKeys("steve.richard32@gmail.com");
                chrome.FindElementByXPath("//*[@id='mobile']").SendKeys("9600674939");
                Thread.Sleep(2000);
                int j = 0;
                while (j < 10)
                {
                    IWebElement captchaTextBox = chrome.FindElement(By.XPath("//*[@id='captcha']"));
                    Thread.Sleep(3000);
                    IJavaScriptExecutor js = (IJavaScriptExecutor)chrome;
                    js.ExecuteScript("arguments[0].scrollIntoView();", captchaTextBox);
                    IWebElement image = chrome.FindElement(By.XPath("//*[@id='imgCaptcha']"));
                    ITakesScreenshot iTScreenshot = (ITakesScreenshot)chrome;
                    Screenshot screenshot1 = iTScreenshot.GetScreenshot();
                    Image imageFromScreen;
                    using (Stream ms = new MemoryStream(screenshot1.AsByteArray))
                    {
                        imageFromScreen = Image.FromStream(ms);
                    }
                    //IWebElement image = driver.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_imgCaptcha']"));
                    float resolutionFactor = imageFromScreen.Width / (float)Screen.PrimaryScreen.Bounds.Width;
                    LogMessage(this.GetType().FullName, "The captcha image point has been identified: " + ((OpenQA.Selenium.Remote.RemoteWebElement)image).LocationOnScreenOnceScrolledIntoView);
                    Point point = ((OpenQA.Selenium.Remote.RemoteWebElement)image).LocationOnScreenOnceScrolledIntoView;
                    int width = (int)Math.Ceiling(image.Size.Width * resolutionFactor);
                    int height = (int)Math.Ceiling(image.Size.Height * resolutionFactor);
                    point.X = (int)Math.Ceiling(point.X * resolutionFactor);
                    point.Y = (int)Math.Ceiling(point.Y * resolutionFactor);
                    Rectangle section = new Rectangle(point, new Size(width, height));
                    Bitmap source = new Bitmap(imageFromScreen);
                    Bitmap finalImage = CropImage(source, section);
                    string pathCropped = Path.Combine(GetWorkingDirectory(), $"GSTRegistrationCaptcha{j}.png");
                    finalImage.Save(pathCropped, ImageFormat.Png);
                    string fileUploadedUrl = uploadFile(pathCropped, botExecutionModel);
                    LogMessage(this.GetType().FullName, "File uploaded successfully");
                    LogMessage(this.GetType().FullName, "The captcha image saved");
                    string captcha = sendRequestAsync(pathCropped, botExecutionModel);
                    LogMessage(this.GetType().FullName, "Captcha " + captcha);
                    chrome.FindElement(By.XPath("//*[@id='captcha']")).SendKeys(captcha);
                    LogMessage(this.GetType().FullName, "Captcha entered in GST portal");
                    Thread.Sleep(4000);
                    chrome.FindElementByXPath("/html/body/div[2]/div[2]/div/div[2]/div/form/div[2]/div/div[2]/div/button").Click();  //submit click
                    LogMessage(this.GetType().FullName, "clicked on submit button");
                    Thread.Sleep(2000);
                    try
                    {
                        chrome.FindElementByXPath("/html/body/div[2]/div[2]/div/div[2]/div/form/div[2]/div/p[1]").Text.Equals("Enter characters as displayed in the CAPTCHA image");
                        LogMessage(this.GetType().FullName, "Captcha Incorrect");
                        Thread.Sleep(3000);
                        j++;
                    }
                    catch (NoSuchElementException)
                    {
                        LogMessage(this.GetType().FullName, "Captcha entered corectly");
                        break;
                    }
                }
                userTask userTask = new userTask();
                List<object> headerValues = new List<object>();
                HeaderValue headerValue = new HeaderValue("Mobile OTP", "text", "Enter the mobile OTP", 20, true, "Please enter the mobile otp", "");
                headerValues.Add(new HeaderValue("Email OTP", "text", "Enter the email OTP", 20, true, "Please enter the email otp", ""));
                headerValues.Add(headerValue);
                AddVariable("GST", headerValues);
                Success("Bot executed successfully");
            }
            catch (Exception e)
            {
                Failure("Bot execution failed " + e.Message);
            }
        }
        private static Bitmap CropImage(Bitmap source, Rectangle section)
        {
            Bitmap bmp = new Bitmap(section.Width, section.Height);
            Graphics g = Graphics.FromImage(bmp);
            g.DrawImage(source, 0, 0, section, GraphicsUnit.Pixel);
            return bmp;
        }
        public string sendRequestAsync(string filePath, BotExecutionModel botExecutionModel)
        {
            //LogMessage(this.GetType().FullName, "Sending request async");
            HttpClient httpClient = new HttpClient();
            MultipartFormDataContent form = new MultipartFormDataContent();
            //LogMessage(this.GetType().FullName, "Sending file path " + filePath);
            byte[] bytes = System.IO.File.ReadAllBytes(filePath);
            string fileName = Path.GetFileName(filePath);
            //LogMessage(this.GetType().FullName, "file name " + fileName);
            string ocrcaptcha = botExecutionModel.proposedBotInputs["ocrcaptcha"].value.ToString();
            LogMessage(this.GetType().FullName, "ocrcaptcha" + ocrcaptcha);
            form.Add(new ByteArrayContent(bytes, 0, bytes.Length), ocrcaptcha, fileName);
            string ocrEndPoint = botExecutionModel.proposedBotInputs["OCREndpoint"].value.ToString();
            //LogMessage(this.GetType().FullName, "OCR Endpoint" + ocrEndPoint);
            HttpResponseMessage response = httpClient.PostAsync(ocrEndPoint, form).Result;
            string stream = response.Content.ReadAsStringAsync().Result;
            dynamic json = JsonConvert.DeserializeObject(stream);
            //LogMessage(this.GetType().FullName, "JSON text" + json);
            string result = json.text;
            //LogMessage(this.GetType().FullName, "Captcha result" + result);
            return result;
        }
        public string uploadFile(string filePath, BotExecutionModel botExecutionModel)
        {
            HttpClient httpClient = new HttpClient();
            httpClient.DefaultRequestHeaders.Authorization =
                new AuthenticationHeaderValue(
                    "Basic", Convert.ToBase64String(
                            System.Text.ASCIIEncoding.ASCII.GetBytes(
                                $"support:password")));
            System.Net.ServicePointManager.ServerCertificateValidationCallback +=
                (se, cert, chain, sslerror) =>
                {
                    return true;
                };
            MultipartFormDataContent form = new MultipartFormDataContent();
            byte[] bytes = System.IO.File.ReadAllBytes(filePath);
            string fileName = Path.GetFileName(filePath);
            string beekeeperDMS = botExecutionModel.proposedBotInputs["BeekeeperDMS"].value.ToString();
            form.Add(new ByteArrayContent(bytes, 0, bytes.Length), "file", fileName);
            HttpResponseMessage response = httpClient.PostAsync(beekeeperDMS, form).Result;
            string stream = response.Content.ReadAsStringAsync().Result;
            dynamic json = JsonConvert.DeserializeObject(stream);
            string result = json.fileDownloadUri;
            return result;
        }
    }

}