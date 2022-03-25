using Beezlabs.RPAHive.Lib;
using Beezlabs.RPAHive.Lib.Models;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security;

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
    public class DetailsOfAuthorisedSignatory
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
        public string documentUploadNameOne { get; set; }
        public string documentUploadNameTwo { get; set; }
    }
    public class BusinessAddress
    {
        public string FlatNo { get; set; }
        public string FloorNo { get; set; }
        public string NameOfBuilding { get; set; }
        public string road { get; set; }
        public string city { get; set; }
        public string district { get; set; }
        public string pincode { get; set; }
        public string officeEmailAddress { get; set; }
        public string officeTelephoneNo { get; set; }
        public string possessionofpremises { get; set; }
        public string proof { get; set; }
        public string natureOfBusiness { get; set; }
        public string AdditionalPaceOfBusiness { get; set; }
    }
    public class GstRegistrationDetailsBot : RPABotTemplate
    {
        List<Details> details = new List<Details>();
        List<DetailsOfPromoter> DetailsOfPromoter = new List<DetailsOfPromoter>();
        List<DetailsOfAuthorisedSignatory> DetailsOfAuthorisedSignatories = new List<DetailsOfAuthorisedSignatory>();
        List<BusinessAddress> businessAddresses = new List<BusinessAddress>();
        protected override void BotLogic(BotExecutionModel botExecutionModel)
        {
            try
            {
                string path = GetWorkingDirectory();
                Details detail = new Details();
                Sharepoint sharepoint = new Sharepoint();
                string webUrl = botExecutionModel.proposedBotInputs["WebUrl"].value.ToString();
                string docLibName = botExecutionModel.proposedBotInputs["DocLibname"].value.ToString();
                string fileName = botExecutionModel.proposedBotInputs["FileName"].value.ToString();
                string IdentityName = botExecutionModel.proposedBotInputs["Credentials"].value.ToString();
                string userName = botExecutionModel.identityList.SingleOrDefault(cred => cred.name.Equals(IdentityName)).credential.basicAuth.username.ToString();
                string password = botExecutionModel.identityList.SingleOrDefault(cred => cred.name.Equals(IdentityName)).credential.basicAuth.password.ToString();
                SecureString secureString = new SecureString();
                string workingDirectory = GetWorkingDirectory().ToString();
                foreach (char c in password.ToCharArray())
                {
                    secureString.AppendChar(c);
                }
                System.Net.ICredentials SPCred = new SharePointOnlineCredentials(userName, secureString);
                sharepoint.DownloadFileViaRestAPI(webUrl, SPCred, docLibName, fileName, workingDirectory);
                using (ExcelFunctions excel = new ExcelFunctions())
                {
                    string excelPath = Path.Combine(path, "GST.xlsx");
                    excel.OpenWorkBook(excelPath);
                    excel.OpenWorkSheet("Details");
                    int row = excel.NumberOfRows();
                    for (int i = 2; i <= row; i++)
                    {
                        detail.iamA = Convert.ToString(excel.GetCellData(i, 1));
                        detail.state = Convert.ToString(excel.GetCellData(i, 2));
                        detail.district = Convert.ToString(excel.GetCellData(i, 3));
                        detail.name = Convert.ToString(excel.GetCellData(i, 4));
                        detail.PAN = Convert.ToString(excel.GetCellData(i, 5));
                        detail.email = Convert.ToString(excel.GetCellData(i, 6));
                        detail.mobileNo = Convert.ToString(excel.GetCellData(i, 7));
                        detail.tradeName = Convert.ToString(excel.GetCellData(i, 8));
                        detail.constOfBusiness = Convert.ToString(excel.GetCellData(i, 9));
                        detail.dateOfCommencement = Convert.ToString(excel.GetCellData(i, 10));
                        detail.typeOfRegistration = Convert.ToString(excel.GetCellData(i, 11));
                        detail.registrationNo = Convert.ToString(excel.GetCellData(i, 12));
                        detail.dateOfRegistration = Convert.ToString(excel.GetCellData(i, 13));
                        details.Add(detail);
                    }
                    excel.OpenWorkSheet("Details-of-Promoter");
                    row = excel.NumberOfRows();
                    for (int j = 2; j <= row; j++)
                    {
                        DetailsOfPromoter detailsOfProprietor = new DetailsOfPromoter();
                        detailsOfProprietor.firstName = Convert.ToString(excel.GetCellData(j, 1));
                        detailsOfProprietor.middleName = Convert.ToString(excel.GetCellData(j, 2));
                        detailsOfProprietor.lastName = Convert.ToString(excel.GetCellData(j, 3));
                        detailsOfProprietor.fatherFirstName = Convert.ToString(excel.GetCellData(j, 4));
                        detailsOfProprietor.fatherMiddleName = Convert.ToString(excel.GetCellData(j, 5));
                        detailsOfProprietor.fatherLastName = Convert.ToString(excel.GetCellData(j, 6));
                        detailsOfProprietor.dob = Convert.ToString(excel.GetCellData(j, 7));
                        detailsOfProprietor.mobileNumber = Convert.ToString(excel.GetCellData(j, 8));
                        detailsOfProprietor.email = Convert.ToString(excel.GetCellData(j, 9));
                        detailsOfProprietor.gender = Convert.ToString(excel.GetCellData(j, 10));
                        detailsOfProprietor.designation = Convert.ToString(excel.GetCellData(j, 11));
                        detailsOfProprietor.DIN = Convert.ToString(excel.GetCellData(j, 12));
                        detailsOfProprietor.PAN = Convert.ToString(excel.GetCellData(j, 13));
                        detailsOfProprietor.flatNo = Convert.ToString(excel.GetCellData(j, 14));
                        detailsOfProprietor.floorNo = Convert.ToString(excel.GetCellData(j, 15));
                        detailsOfProprietor.nameofBuilding = Convert.ToString(excel.GetCellData(j, 16));
                        detailsOfProprietor.Street = Convert.ToString(excel.GetCellData(j, 17));
                        detailsOfProprietor.city = Convert.ToString(excel.GetCellData(j, 18));
                        detailsOfProprietor.state = Convert.ToString(excel.GetCellData(j, 19));
                        detailsOfProprietor.district = Convert.ToString(excel.GetCellData(j, 20));
                        detailsOfProprietor.zipCode = Convert.ToString(excel.GetCellData(j, 21));
                        detailsOfProprietor.documentUploadName = Convert.ToString(excel.GetCellData(j, 22));
                        DetailsOfPromoter.Add(detailsOfProprietor);
                    }
                    excel.OpenWorkSheet("Details-of-Authorized-Signatory");
                    row = excel.NumberOfRows();
                    for (int k = 2; k <= row; k++)
                    {
                        DetailsOfAuthorisedSignatory detailsOfAuthorisedSignatory = new DetailsOfAuthorisedSignatory();
                        detailsOfAuthorisedSignatory.firstName = Convert.ToString(excel.GetCellData(k, 1));
                        detailsOfAuthorisedSignatory.middleName = Convert.ToString(excel.GetCellData(k, 2));
                        detailsOfAuthorisedSignatory.lastName = Convert.ToString(excel.GetCellData(k, 3));
                        detailsOfAuthorisedSignatory.fatherFirstName = Convert.ToString(excel.GetCellData(k, 4));
                        detailsOfAuthorisedSignatory.fatherMiddleName = Convert.ToString(excel.GetCellData(k, 5));
                        detailsOfAuthorisedSignatory.fatherLastName = Convert.ToString(excel.GetCellData(k, 6));
                        detailsOfAuthorisedSignatory.dob = Convert.ToString(excel.GetCellData(k, 7));
                        detailsOfAuthorisedSignatory.mobileNumber = Convert.ToString(excel.GetCellData(k, 8));
                        detailsOfAuthorisedSignatory.email = Convert.ToString(excel.GetCellData(k, 9));
                        detailsOfAuthorisedSignatory.gender = Convert.ToString(excel.GetCellData(k, 10));
                        detailsOfAuthorisedSignatory.designation = Convert.ToString(excel.GetCellData(k, 11));
                        detailsOfAuthorisedSignatory.DIN = Convert.ToString(excel.GetCellData(k, 12));
                        detailsOfAuthorisedSignatory.PAN = Convert.ToString(excel.GetCellData(k, 13));
                        detailsOfAuthorisedSignatory.flatNo = Convert.ToString(excel.GetCellData(k, 14));
                        detailsOfAuthorisedSignatory.floorNo = Convert.ToString(excel.GetCellData(k, 15));
                        detailsOfAuthorisedSignatory.nameofBuilding = Convert.ToString(excel.GetCellData(k, 16));
                        detailsOfAuthorisedSignatory.Street = Convert.ToString(excel.GetCellData(k, 17));
                        detailsOfAuthorisedSignatory.city = Convert.ToString(excel.GetCellData(k, 18));
                        detailsOfAuthorisedSignatory.state = Convert.ToString(excel.GetCellData(k, 19));
                        detailsOfAuthorisedSignatory.district = Convert.ToString(excel.GetCellData(k, 20));
                        detailsOfAuthorisedSignatory.zipCode = Convert.ToString(excel.GetCellData(k, 21));
                        detailsOfAuthorisedSignatory.documentUploadNameOne = Convert.ToString(excel.GetCellData(k, 22));
                        detailsOfAuthorisedSignatory.documentUploadNameTwo = Convert.ToString(excel.GetCellData(k, 23));
                        DetailsOfAuthorisedSignatories.Add(detailsOfAuthorisedSignatory);
                    }
                    excel.OpenWorkSheet("Principal-Place-of-Business");
                    row = excel.NumberOfRows();
                    for (int l = 2; l <= row; l++)
                    {
                        BusinessAddress businessAddress = new BusinessAddress();
                        businessAddress.FlatNo = Convert.ToString(excel.GetCellData(l, 1));
                        businessAddress.FloorNo = Convert.ToString(excel.GetCellData(l, 2));
                        businessAddress.NameOfBuilding = Convert.ToString(excel.GetCellData(l, 3));
                        businessAddress.road = Convert.ToString(excel.GetCellData(l, 4));
                        businessAddress.city = Convert.ToString(excel.GetCellData(l, 5));
                        businessAddress.district = Convert.ToString(excel.GetCellData(l, 6));
                        businessAddress.pincode = Convert.ToString(excel.GetCellData(l, 7));
                        businessAddress.officeEmailAddress = Convert.ToString(excel.GetCellData(l, 8));
                        businessAddress.officeTelephoneNo = Convert.ToString(excel.GetCellData(l, 9));
                        businessAddress.natureOfBusiness = Convert.ToString(excel.GetCellData(l, 10));
                        businessAddress.AdditionalPaceOfBusiness = Convert.ToString(excel.GetCellData(l, 11));
                        businessAddress.proof = Convert.ToString(excel.GetCellData(l, 12));
                        businessAddresses.Add(businessAddress);
                    }
                }
                AddVariable("details", details);
                AddVariable("DetailsOfPromoter", DetailsOfPromoter);
                AddVariable("DetailsOfAuthorisedSignatories", DetailsOfAuthorisedSignatories);
                AddVariable("businessAddresses", businessAddresses);
                Success("Bot executed successfully");
            }
            catch (Exception e)
            {
                Failure("Bot execution failed due to " + e.Message);
            }
        }
    }
}