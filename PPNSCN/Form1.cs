using Newtonsoft.Json;
using PPNSCN.Model;
using PPNSCN.Properties;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace PPNSCN
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.statusBar.Location = new Point(387, 466);
            this.statusBar.Size = new System.Drawing.Size(217, 92);
            this.statusBar.Image = Resources.guide;
            this.statusBar.Refresh();
            System.Threading.Thread.Sleep(1000);
        }

        private void fetchdata_Click(object sender, EventArgs e)
        {

            try
            {
                this.statusBar.Location = new Point(387, 466);
                this.statusBar.Size = new System.Drawing.Size(82, 92);
                this.statusBar.Image = Resources.wait;
                this.statusBar.Refresh();
                System.Threading.Thread.Sleep(3000);

                var directory = new DirectoryInfo("C:\\HajOnSoft");
                var pphoto = directory.GetFiles().Where(x => x.Name.Contains("IMAGEPHOTO"))
             .OrderByDescending(f => f.LastWriteTime)
             .First();
                personphoto.Image = GetCopyOfImage(pphoto.FullName);

                var passportpic = directory.GetFiles().Where(x => x.Name.Contains("IMAGEVIS"))
             .OrderByDescending(f => f.LastWriteTime)
             .First();

                passportphotos.Image = GetCopyOfImage(passportpic.FullName);


                var MRZ = directory.GetFiles().Where(x => x.Name.Contains("CODELINE"))
                                .OrderByDescending(f => f.LastWriteTime)
                                .First();
                string text = File.ReadAllText(MRZ.FullName);
                if (text.Length > 80)
                {
                    string algo = "Ptiiinnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnn#########CbbbYYMMDDCsyymmddCppppppppppppppCX";


                    // var NameArray = text.Substring(algo.IndexOf('n'), algo.Count(x => x == 'n')).Split('<').Where(x => x != "").ToArray();
                    var DocumentTypeArray = text.Substring(algo.IndexOf('P'), algo.Count(x => x == 'P'));
                    var IssueCountryArray = text.Substring(algo.IndexOf('i'), algo.Count(x => x == 'i'));
                    var PassportNumber = Regex.Replace(text.Substring(algo.IndexOf('#'), algo.Count(x => x == '#')), @"\t|\n|\r", "");
                    var Nationality = text.Substring(algo.IndexOf('b'), 4);
                    var Gender = text.Substring(20 + 45, 1);

                    this.passportnumber.Text = PassportNumber;
                    //this.firstname.Text = NameArray[1];
                    //this.lastname.Text = NameArray[0];

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
                    this.statusBar.Location = new Point(387, 460);
                    this.statusBar.Size = new System.Drawing.Size(120, 110);
                    this.statusBar.Image = Resources.success;
                    this.statusBar.Refresh();
                    System.Threading.Thread.Sleep(1000);
                }
                else
                {
                    this.statusBar.Location = new Point(387, 460);
                    this.statusBar.Size = new System.Drawing.Size(120, 110);
                    this.statusBar.Image = Resources.cross;
                    this.statusBar.Refresh();
                    personphoto.Image = null;
                    this.personphoto.Refresh();
                    passportphotos.Image = null;
                    this.passportphotos.Refresh();

                    this.passportnumber.Text = "";
                    this.firstname.Text = "";
                    this.lastname.Text = "";
                    this.dob.Text = "";
                    this.expirydate.Text = "";
                    this.gender.Text = "";


                    System.Threading.Thread.Sleep(1000);

                }
            }
            catch (Exception ex)
            {
                this.statusBar.Location = new Point(387, 460);
                this.statusBar.Size = new System.Drawing.Size(120, 110);
                this.statusBar.Image = Resources.cross;
                this.statusBar.Refresh();
                personphoto.Image = null;
                this.personphoto.Refresh();
                passportphotos.Image = null;
                this.passportphotos.Refresh();

                this.passportnumber.Text = "";
                this.firstname.Text = "";
                this.lastname.Text = "";
                this.dob.Text = "";
                this.expirydate.Text = "";
                this.gender.Text = "";


                System.Threading.Thread.Sleep(1000);
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

        private void GetSignature_Click(object sender, EventArgs e)
        {
            var directory = new DirectoryInfo("C:\\HajOnSoft");
            var signaturepic = directory.GetFiles().Where(x => x.Name.Contains("Signature"))
                .OrderByDescending(f => f.LastWriteTime)
                .First();

            SignatureBox.Image = GetCopyOfImage(signaturepic.FullName);

        }
    }
}
