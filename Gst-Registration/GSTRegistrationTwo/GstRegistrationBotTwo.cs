using Beezlabs.RPAHive.Lib;
using Beezlabs.RPAHive.Lib.Models;
using Newtonsoft.Json;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
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
    public class GstRegistrationBotTwo : RPABotTemplate
    {
        protected override void BotLogic(BotExecutionModel botExecutionModel)
        {
            try
            {
                ChromeOptions options = new ChromeOptions();
                options.BinaryLocation = "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe";
                options.DebuggerAddress = "localhost:9876";
                ChromeDriver chrome = new ChromeDriver(GetBotPath(), options);
                string emailOtp = botExecutionModel.proposedBotInputs["Email OTP"].value.ToString();
                string mobileOtp = botExecutionModel.proposedBotInputs["Mobile OTP"].value.ToString();
                chrome.FindElementByXPath("//*[@id='mobile_otp']").SendKeys(mobileOtp);
                LogMessage(this.GetType().FullName, "Fetched mobile otp");
                chrome.FindElementByXPath("//*[@id='email-otp']").SendKeys(emailOtp);
                LogMessage(this.GetType().FullName, "fetched email otp");
                chrome.FindElementByXPath("/html/body/div[2]/div[2]/div/div[2]/div/div[2]/div/form/div/div/button").Click();
                LogMessage(this.GetType().FullName, "Clicked on proceed button");
                Thread.Sleep(5000);
                string trn = chrome.FindElementByXPath("/html/body/div[2]/div[2]/div/div[2]/div/div[2]/div/div[1]/div/p/span[2]").Text;
                LogMessage(this.GetType().FullName, "trn " +trn);
                chrome.FindElementByXPath("/html/body/div[2]/div[2]/div/div[2]/div/div[2]/div/div[2]/div/a").Click();
                LogMessage(this.GetType().FullName, "clicked on trn proceed");
                chrome.FindElementByXPath("//*[@id='trnno']").SendKeys(trn);
                int j = 0;
                while (j < 10)
                {
                    IWebElement captchaTextBox = chrome.FindElement(By.XPath("//*[@id='captchatrn']"));
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
                    chrome.FindElement(By.XPath("//*[@id='captchatrn']")).SendKeys(captcha);
                    LogMessage(this.GetType().FullName, "Captcha entered in GST portal");
                    Thread.Sleep(4000);
                    chrome.FindElementByXPath("/html/body/div[2]/div[2]/div/div[2]/div/form/div[2]/div/div[2]/div/button").Click();  //submit click
                    LogMessage(this.GetType().FullName, "clicked on submit button");
                    Thread.Sleep(2000);
                    try
                    {
                        chrome.FindElementByXPath("/html/body/div[2]/div[2]/div/div[2]/div/form/div[2]/div/fieldset/div/div[2]/div/div/span").Text.Equals("Enter characters as displayed in the CAPTCHA image");
                        LogMessage(this.GetType().FullName, "Captcha Incorrect");
                        Thread.Sleep(3000);
                        j++;
                    }
                    catch (NoSuchElementException ex)
                    {
                        LogMessage(this.GetType().FullName, "Captcha entered corectly");
                        break;
                    }
                }                
                userTask userTask = new userTask();
                List<object> headerValues = new List<object>();
                HeaderValue headerValue = new HeaderValue("SecondMobileOTP", "text", "Enter the mobile OTP", 20, true, "Please enter the mobile otp", "");
                headerValues.Add(headerValue);
                AddVariable("GSTRegistrationTaskTwo", headerValues);
                AddVariable("TRN", trn);
                Success("Bot executed successfully");
            }
            catch(Exception e)
            {
                Failure("Bot execution failed due to " + e.Message);
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