using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using PPNSCN.Model;
using PPNSCN.Properties;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace PPNSCN
{
    public partial class Form1 : Form
    {
        public class SaveDataModel
        {
            public string DocumentNumber { get; set; }
            public string FirstName { get; set; }
            public string LastName { get; set; }
            public string DOB { get; set; }
            public string Gender { get; set; }
            public string IssuingCountry { get; set; }
            public string Nationality { get; set; }
            public string ExpiryDate { get; set; }
            public List<ImageAttachment> ImageAttachments { get; set; }
        }

        public class ImageAttachment
        {
            public string FileName { get; set; }
            public string Attachment { get; set; }
            public string DocumentType { get; set; }
            public string Page { get; set; }
        }
        private Image PassportPerson { get; set; }
        private Image PassportPage1 { get; set; }
        private Image PassportPage2 { get; set; }
        private Image EmiratedIdFront { get; set; }
        private Image EmiratedIdBack { get; set; }
        private Image ResidencePermit { get; set; }
        private Image DrivingLicenseFront { get; set; }
        private Image DrivingLicenseBack { get; set; }
        private Image SignatureImage { get; set; }

        private string ImageSaveDestination = ConfigurationManager.AppSettings["DataSaveToLocal"];

        private string ApiURl = ConfigurationManager.AppSettings["ApiURl"];

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

            //this.statusBar.Location = new Point(387, 466);
            //this.statusBar.Size = new System.Drawing.Size(217, 92);
            //this.statusBar.Image = Resources.guide;
            //this.statusBar.Refresh();
            System.Threading.Thread.Sleep(1000);
        }

        private void fetchdata_Click()
        {
            try
            {
                var directory = new DirectoryInfo("C:\\HajOnSoft");
                var MRZ = directory.GetFiles().Where(x => x.Name.Contains("CODELINE"))
                                .OrderByDescending(f => f.LastWriteTime)
                                .First();
                string text = File.ReadAllText(MRZ.FullName);
                if (text.Length > 80)
                {
                    string algo = "Ptiiinnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnn#########CbbbYYMMDDCsyymmddCppppppppppppppCX";
                    // var NameArray = text.Substring(algo.IndexOf('n'), algo.Count(x => x == 'n')).Split('<').Where(x => x != "").ToArray();
                    var DocumentTypeArray = text.Substring(algo.IndexOf('P'), algo.Count(x => x == 'P'));
                    var IssueCountryArray = text.Substring(2, 3);
                    var PassportNumber = Regex.Replace(text.Substring(algo.IndexOf('#'), algo.Count(x => x == '#')), @"\t|\n|\r", "");
                    var Nationality = text.Substring(algo.IndexOf('b'), 4);
                    var Gender = text.Substring(20 + 45, 1);
                    this.passportnumber.Text = PassportNumber;

                    var nameArraySplit = text.Substring(5).Split(new[] { "<<" }, StringSplitOptions.RemoveEmptyEntries);
                    this.firstname.Text = nameArraySplit.Length >= 2 ? nameArraySplit[1].Replace("<", " ") : nameArraySplit[0].Replace("<", " ");
                    var nameArraySplitLast = text.Substring(5).Split(new[] { "<<" }, StringSplitOptions.RemoveEmptyEntries);
                    this.lastname.Text = nameArraySplit.Length >= 2 ? nameArraySplitLast[0].Replace("<", " ") : string.Empty;
                    this.dob.Text = DateOfBirth(text).ToString("dd/MM/yyyy");
                    this.expirydate.Text = ExpireDate(text).ToString("dd/MM/yyyy");
                    this.gender.Text = Gender.Equals("M") ? "Male" : Gender.Equals("F") ? "Female" : "Other";
                    var json = System.Text.Encoding.UTF8.GetString(Resources.ISOCountryCode);
                    var GetParseJsonArray = JsonConvert.DeserializeObject<List<JsonParserISOCountry>>(json).Where(x => x.threecode.Equals(Regex.Replace(Nationality, @"[\d-]", string.Empty))).FirstOrDefault();
                    this.nationality.Text = GetParseJsonArray != null ? GetParseJsonArray.Key : Regex.Replace(Nationality, @"[\d-]", string.Empty);
                    var GetParseIssueJsonArray = JsonConvert.DeserializeObject<List<JsonParserISOCountry>>(json).Where(x => x.threecode.Equals(Regex.Replace(IssueCountryArray, @"[\d-]", string.Empty))).FirstOrDefault();
                    this.IssueCountry.Text = GetParseIssueJsonArray != null ? GetParseIssueJsonArray.Key : Regex.Replace(IssueCountryArray, @"[\d-]", string.Empty);

                }
                else
                {
                    this.passportnumber.Text = "";
                    this.firstname.Text = "";
                    this.lastname.Text = "";
                    this.dob.Text = "";
                    this.expirydate.Text = "";
                    this.gender.Text = "";
                    this.IssueCountry.Text = "";
                    this.nationality.Text = "";
                }
            }
            catch (Exception ex)
            {
                this.passportnumber.Text = "";
                this.firstname.Text = "";
                this.lastname.Text = "";
                this.dob.Text = "";
                this.expirydate.Text = "";
                this.gender.Text = "";
                this.IssueCountry.Text = "";
                this.nationality.Text = "";
                MessageBox.Show(ex.Message, "Error");
            }


        } 

        private void ScanApp_Click(object sender, EventArgs e)
        {
            string workingDirectory = Environment.CurrentDirectory;
            //  string startupPath = Directory.GetParent(workingDirectory).Parent.FullName;
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.FileName = workingDirectory + "\\Scan\\OCR640.exe";
            startInfo.Arguments = "header.h";
            //  startInfo.WindowStyle = ProcessWindowStyle.Hidden;
            Process.Start(startInfo);
        }

        private void SignationApp(object sender, EventArgs e)
        {
            string workingDirectory = Environment.CurrentDirectory;
            //  string startupPath = Directory.GetParent(workingDirectory).Parent.FullName;
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.FileName = workingDirectory + "\\Signature\\Signature_Pad_By_Onpro.exe";
            startInfo.Arguments = "header.h";
            //  startInfo.WindowStyle = ProcessWindowStyle.Hidden;
            Process.Start(startInfo);
        }

        private Image GetCopyOfImage(string path)
        {
            Image image;
            using (FileStream myStream = new FileStream(path, FileMode.Open))
            {
                image = Image.FromStream(myStream);
            }

            return image;
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void saveDataBtn_Click(object sender, EventArgs e)
        {
            List<ImageAttachment> AttachmentData = new List<ImageAttachment>();
            if (PassportPerson != null)
                SaveAs(PassportPerson, ImageSaveDestination + this.passportnumber.Text, "Face.jpg", ref AttachmentData, "Face", "passport");
            if (PassportPage1 != null)
                SaveAs(PassportPage1, ImageSaveDestination + this.passportnumber.Text, "PassportFrontPage.jpg", ref AttachmentData, "Page1", "passport");
            if (PassportPage2 != null)
                SaveAs(PassportPage2, ImageSaveDestination + this.passportnumber.Text, "PassportBackPage.jpg", ref AttachmentData, "Page2", "passport");
            if (EmiratedIdFront != null)
                SaveAs(EmiratedIdFront, ImageSaveDestination + this.passportnumber.Text, "EmiratesIdFrontPage.jpg", ref AttachmentData, "Front", "emirates_id");
            if (EmiratedIdBack != null)
                SaveAs(EmiratedIdBack, ImageSaveDestination + this.passportnumber.Text, "EmiratesIdBackPage.jpg", ref AttachmentData, "Back", "emirates_id");
            if (ResidencePermit != null)
                SaveAs(ResidencePermit, ImageSaveDestination + this.passportnumber.Text, "ResidencePermit.jpg", ref AttachmentData, "Front", "residence_permit");
            if (DrivingLicenseFront != null)
                SaveAs(DrivingLicenseFront, ImageSaveDestination + this.passportnumber.Text, "DrivingLicenseFront.jpg", ref AttachmentData, "Front", "driving_license");
            if (DrivingLicenseBack != null)
                SaveAs(DrivingLicenseBack, ImageSaveDestination + this.passportnumber.Text, "DrivingLicenseBack.jpg", ref AttachmentData, "Back", "driving_license");
            if (SignatureImage != null)
                SaveAs(SignatureImage, ImageSaveDestination + this.passportnumber.Text, "SignatureImage.jpg", ref AttachmentData, "Page1", "signature");

            if (PostData(ApiURl, AttachmentData))
                MessageBox.Show("Successfully Saved!", "Alert");
            else
                MessageBox.Show("Error While Saving Data", "Alert");
        }

        private void SaveCurrentDocument_Click(object sender, EventArgs e)
        {
            try
            {
                var directory = new DirectoryInfo("C:\\HajOnSoft");
                switch (this.DocumentCombo.SelectedItem)
                {
                    case "Signature":
                        var signaturepic = directory.GetFiles().Where(x => x.Name.Contains("Signature"))
                            .OrderByDescending(f => f.LastWriteTime)
                            .First();
                        SignatureImage = GetCopyOfImage(signaturepic.FullName);
                        signaturebox.Image = SignatureImage;
                        break;

                    case "Passport":
                        switch (this.DocumentTypeCombo.SelectedItem)
                        {
                            case "Page 1":
                                this.pp1.BackColor = Color.Green;
                                var passportpic = directory.GetFiles().Where(x => x.Name.Contains("IMAGEVIS"))
                                                   .OrderByDescending(f => f.LastWriteTime)
                                                   .First();
                                PassportPage1 = GetCopyOfImage(passportpic.FullName);

                                var pphoto = directory.GetFiles().Where(x => x.Name.Contains("IMAGEPHOTO"))
                                                 .OrderByDescending(f => f.LastWriteTime)
                                                 .First();
                                PassportPerson = GetCopyOfImage(pphoto.FullName);
                                personphoto.Image = PassportPerson;
                                passportphotos.Image = PassportPage1;
                                fetchdata_Click();
                                break;


                            case "Page 2":
                                this.pp2.BackColor = Color.Green;
                                var passportpic2 = directory.GetFiles().Where(x => x.Name.Contains("IMAGEVIS"))
                                                   .OrderByDescending(f => f.LastWriteTime)
                                                   .First();
                                PassportPage2 = GetCopyOfImage(passportpic2.FullName);
                                passportphoto2.Image = PassportPage2;
                                break;

                            default:
                                this.pp1.BackColor = Color.Red;
                                this.pp2.BackColor = Color.Red;
                                MessageBox.Show("Select Different Document Type", "Error");
                                break;
                        }
                        break;


                    case "Emirates ID":
                        switch (this.DocumentTypeCombo.SelectedItem)
                        {
                            case "Front":
                                this.eid1.BackColor = Color.Green;
                                var emiratesid1 = directory.GetFiles().Where(x => x.Name.Contains("IMAGEVIS"))
                                                .OrderByDescending(f => f.LastWriteTime)
                                                .First();
                                EmiratedIdFront = GetCopyOfImage(emiratesid1.FullName);
                                EIDFront.Image = EmiratedIdFront;
                                break;


                            case "Back":
                                this.eid2.BackColor = Color.Green;
                                var emiratesid2 = directory.GetFiles().Where(x => x.Name.Contains("IMAGEVIS"))
                                             .OrderByDescending(f => f.LastWriteTime)
                                             .First();
                                EmiratedIdBack = GetCopyOfImage(emiratesid2.FullName);
                                EIDBack.Image = EmiratedIdBack;
                                break;

                            default:
                                this.eid1.BackColor = Color.Red;
                                this.eid2.BackColor = Color.Red;
                                MessageBox.Show("Select Different Document Type", "Error");
                                break;
                        }
                        break;


                    case "UAE Residence Visa":
                        switch (this.DocumentTypeCombo.SelectedItem)
                        {
                            case "Page 1":
                                this.rp1.BackColor = Color.Green;
                                var Rpermit1 = directory.GetFiles().Where(x => x.Name.Contains("IMAGEVIS"))
                                   .OrderByDescending(f => f.LastWriteTime)
                                   .First();
                                ResidencePermit = GetCopyOfImage(Rpermit1.FullName);
                                RPBox.Image = ResidencePermit;
                                break;

                            default:
                                this.rp1.BackColor = Color.Red;
                                MessageBox.Show("Select Different Document Type", "Error");
                                break;
                        }

                        break;

                    case "UAE Driving Licence":
                        switch (this.DocumentTypeCombo.SelectedItem)
                        {
                            case "Front":
                                this.dl1.BackColor = Color.Green;
                                var Drivinglicense1 = directory.GetFiles().Where(x => x.Name.Contains("IMAGEVIS"))
                                                               .OrderByDescending(f => f.LastWriteTime)
                                                               .First();
                                DrivingLicenseFront = GetCopyOfImage(Drivinglicense1.FullName);
                                DrivingLicense1.Image = DrivingLicenseFront;
                                break;


                            case "Back":
                                this.dl2.BackColor = Color.Green;
                                var Drivinglicense2 = directory.GetFiles().Where(x => x.Name.Contains("IMAGEVIS"))
                                                      .OrderByDescending(f => f.LastWriteTime)
                                                      .First();
                                DrivingLicenseBack = GetCopyOfImage(Drivinglicense2.FullName);
                                DrivingLicense2.Image = DrivingLicenseBack;
                                break;

                            default:
                                this.dl1.BackColor = Color.Red;
                                this.dl2.BackColor = Color.Red;
                                MessageBox.Show("Select Different Document Type", "Error");
                                break;
                        }
                        break;

                    default:
                        MessageBox.Show("Select Different Document Type", "Error");
                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Please Re-scan This Document", "Error");
            }
        }

        private void comboBox1_TextChanged(object sender, EventArgs e)
        {
            switch (((System.Windows.Forms.ComboBox)sender).SelectedItem)
            {
                case "Passport":
                    this.DocumentTypeCombo.Text = "";
                    this.DocumentTypeCombo.Items.Clear();
                    this.DocumentTypeCombo.Items.AddRange(new object[] { "Page 1", "Page 2" });
                    this.signaturebtn.Visible = false;
                    break;

                case "Signature":
                    this.DocumentTypeCombo.Text = "";
                    this.DocumentTypeCombo.Items.Clear();
                    this.DocumentTypeCombo.Items.AddRange(new object[] { "Signature" });
                    this.signaturebtn.Visible = true;
                    break;

                case "Emirates ID":
                    this.DocumentTypeCombo.Text = "";
                    this.DocumentTypeCombo.Items.Clear();
                    this.DocumentTypeCombo.Items.AddRange(new object[] { "Front", "Back" });
                    this.signaturebtn.Visible = false;
                    break;


                case "UAE Residence Visa":
                    this.DocumentTypeCombo.Text = "";
                    this.DocumentTypeCombo.Items.Clear();
                    this.DocumentTypeCombo.Items.AddRange(new object[] { "Page 1" });
                    this.signaturebtn.Visible = false;
                    break;

                case "UAE Driving Licence":
                    this.DocumentTypeCombo.Text = "";
                    this.DocumentTypeCombo.Items.Clear();
                    this.DocumentTypeCombo.Items.AddRange(new object[] { "Front", "Back" });
                    this.signaturebtn.Visible = false;
                    break;

                default:
                    MessageBox.Show("Select Different Document", "Error");
                    break;
            }
        }

        private void ResetDataBtn_Click(object sender, EventArgs e)
        {
            PassportPerson = null;
            PassportPage1 = null;
            PassportPage2 = null;
            EmiratedIdFront = null;
            EmiratedIdBack = null;
            ResidencePermit = null;
            DrivingLicenseFront = null;
            DrivingLicenseBack = null;
            this.passportnumber.Text = "";
            this.firstname.Text = "";
            this.lastname.Text = "";
            this.dob.Text = "";
            this.expirydate.Text = "";
            this.gender.Text = "";
            this.IssueCountry.Text = "";
            this.nationality.Text = "";


            personphoto.Image = PassportPerson;
            passportphotos.Image = PassportPage1;
            passportphoto2.Image = PassportPage2;
            EIDFront.Image = EmiratedIdFront;
            EIDBack.Image = EmiratedIdBack;
            RPBox.Image = ResidencePermit;
            DrivingLicense1.Image = DrivingLicenseFront;
            DrivingLicense2.Image = DrivingLicenseBack;



            this.pp1.BackColor = Color.Red;
            this.pp2.BackColor = Color.Red;
            this.eid1.BackColor = Color.Red;
            this.eid2.BackColor = Color.Red;
            this.rp1.BackColor = Color.Red;
            this.dl1.BackColor = Color.Red;
            this.dl2.BackColor = Color.Red;
        }

        private void SaveAs(Image FileUpload, string appPath, string Filename, ref List<ImageAttachment> AttachmentData, string page, string doctype)
        {
            if (!Directory.Exists(appPath))
                Directory.CreateDirectory(appPath);
            if (System.IO.File.Exists(appPath + "\\" + Filename))
                System.IO.File.Delete(appPath + "\\" + Filename);
            new Bitmap(FileUpload).Save(appPath + "\\" + Filename, ImageFormat.Jpeg);
            AttachmentData.Add(new ImageAttachment() { Attachment = GetBase64(FileUpload), FileName = Filename, DocumentType = doctype, Page = page });
        }

        private bool PostData(string apiurl, List<ImageAttachment> AttachmentData)
        {
            bool data = false;
            try
            {
                HttpWebRequest webrequest = (HttpWebRequest)WebRequest.Create(apiurl);
                webrequest.Method = "POST";
                webrequest.ContentType = "application/json";
                webrequest.ContentLength = 0;
                var datatosave = new SaveDataModel()
                {
                    DOB = dob.Text,
                    DocumentNumber = passportnumber.Text,
                    ExpiryDate = expirydate.Text,
                    IssuingCountry = IssueCountry.Text,
                    FirstName = firstname.Text,
                    LastName = lastname.Text,
                    Gender = gender.Text,
                    Nationality = nationality.Text,
                    ImageAttachments = AttachmentData
                };
                var json = JsonConvert.SerializeObject(datatosave);
                webrequest.ContentLength = json.Length;
                using (var streamWriter = new StreamWriter(webrequest.GetRequestStream()))
                {
                    streamWriter.Write(json);
                }
                HttpWebResponse webresponse = (HttpWebResponse)webrequest.GetResponse();
                Encoding enc = System.Text.Encoding.GetEncoding("utf-8");
                StreamReader responseStream = new StreamReader(webresponse.GetResponseStream(), enc);
                string result = string.Empty;
                result = responseStream.ReadToEnd();
                webresponse.Close();
                data = (bool)JsonConvert.DeserializeObject(result);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Saving Error!");
                return false;
            }
            return data;
        }

        public string GetBase64(Image image)
        {
            using (MemoryStream m = new MemoryStream())
            {
                var i = image;

                var i2 = new Bitmap(i);
                i2.Save(m, ImageFormat.Jpeg);
                byte[] imageBytes = m.ToArray();
                string base64String = Convert.ToBase64String(imageBytes);
                return base64String;
            }
        }

        private DateTime DateOfBirth(string mrz)
        {
            var dob = new DateTime(int.Parse(DateTime.Now.Year.ToString().Substring(0, 2) + mrz.Substring(14 + 44, 2)), int.Parse(mrz.Substring(16 + 44, 2)),
                    int.Parse(mrz.Substring(18 + 44, 2)));

            if (dob < DateTime.Now)
                return dob;

            return dob.AddYears(-100); //Subtract a century

        }

        private DateTime ExpireDate(string mrz)
        {
            //I am assuming all passports will certainly expire this century
            return new DateTime(int.Parse(DateTime.Now.Year.ToString().Substring(0, 2) + mrz.Substring(22 + 44, 2)), int.Parse(mrz.Substring(24 + 44, 2)),
                int.Parse(mrz.Substring(26 + 44, 2)));
        }

    }
}
