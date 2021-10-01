using DevExpress.XtraEditors;
using DevExpress.XtraReports.Design;
using DevExpress.XtraSplashScreen;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace TD_PrintInvoice
{
    public partial class Form1 : DevExpress.XtraEditors.XtraForm
    {
        Global fnc = new Global();
        private string outputPdfPath = @"C:\PrintInvoice";
        private string[] invoiceRefGroup;
        private string[] einvoiceGroup;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            try
            {
                fnc.DBConnection();
                string[] args = Environment.GetCommandLineArgs();
                fnc.LogoArg(args);
                invoiceRefGroup = fnc.invoiceRef.Split(',');
                einvoiceGroup = fnc.einvoice.Split(',');
                if (invoiceRefGroup.Length < 1)
                {
                    return;
                }
                GetInvoicePDF(invoiceRefGroup, einvoiceGroup);
                GroupPDF(invoiceRefGroup,fnc.userName);
                axAcroPDF1.src = outputPdfPath;
            }
            catch(Exception ex)
            {
                XtraMessageBox.Show(ex.Message.ToString());
            }
            SplashScreenManager.CloseForm();
        }
        public void GetInvoicePDF(string[] invoiceRefGroup, string[] einvoiceGroup)
        {
            try
            {
                LDXCComApi.Client LDXClient = new LDXCComApi.Client();
                LDXClient.Login(fnc.connectUsername, fnc.connectPassword, 18);
                if (LDXClient.IsLoggedIn)
                {
                    for (int i = 0; i < invoiceRefGroup.Length; i++)
                    {
                        string path = "C:\\PrintInvoice";
                        var fatura = (dynamic)null;
                        LDXCComApi.EInvoice eInvoice = LDXClient.CreateEInvoice();
                        if (Convert.ToInt32(einvoiceGroup[i]) == 1)
                        {
                            fatura = eInvoice.AddUnityInvoice(Convert.ToInt32(invoiceRefGroup[i]), LDXCComApi.EInvoiceTypes.etSales);
                        }
                        else
                        {
                            fatura = eInvoice.AddEArchiveInvoice(Convert.ToInt32(invoiceRefGroup[i]));
                        }
                        if (fatura.GUID != null)
                        {
                            LDXCComApi.EInvoiceElement inv = eInvoice.GetByETTN(fatura.GUID);
                            path = string.Concat(path, "\\", invoiceRefGroup[i], ".pdf");
                            inv.SavePDFToFile(path);
                        }
                        else
                        {
                            XtraMessageBox.Show(fatura.ErrorInfo);
                        }
                    }
                    LDXClient.Logout();
                }
                else
                {
                    XtraMessageBox.Show("LogoConnect Bağlantısı Kurulamadı!.Hata: " + LDXClient.ErrorCode.ToString() + "," + LDXClient.ErrorInfo);
                }

            }
            catch (Exception ex)
            {
                XtraMessageBox.Show("LogoConnect Bağlantısı Kurulurken Hata Oluştu: " + ex.Message);
            }
        }

        private void GroupPDF(string[] invoiceRefGroup, string userName)
        {
            outputPdfPath = string.Concat(outputPdfPath,"\\", userName, "_dispatch.pdf");
            PdfReader reader = null;
            Document sourceDocument = null;
            PdfCopy pdfCopyProvider = null;
            PdfImportedPage importedPage;
            sourceDocument = new Document();
            pdfCopyProvider = new PdfCopy(sourceDocument, new System.IO.FileStream(outputPdfPath, System.IO.FileMode.Create));
            sourceDocument.Open();
            try
            {
                for (int f = 0; f < invoiceRefGroup.Length; f++)
                {
                    string path = "C:\\PrintInvoice";
                    path = string.Concat(path, "\\", invoiceRefGroup[f], ".pdf");
                    int pages = get_pageCcount(path);
                    reader = new PdfReader(path);
                    for (int i = 1; i <= pages; i++)
                    {
                        importedPage = pdfCopyProvider.GetImportedPage(reader, i);
                        pdfCopyProvider.AddPage(importedPage);
                    }

                    reader.Close();
                }
                sourceDocument.Close();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        private int get_pageCcount(string file)
        {
            using (StreamReader sr = new StreamReader(File.OpenRead(file)))
            {
                Regex regex = new Regex(@"/Type\s*/Page[^s]");
                MatchCollection matches = regex.Matches(sr.ReadToEnd());

                return matches.Count;
            }
        }

        private void SimpleButton1_Click(object sender, EventArgs e)
        {
            
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            DeletePDF(invoiceRefGroup);
        }
        private void DeletePDF(string[] invoiceRefGroup)
        {
            for (int i = 0; i < invoiceRefGroup.Length; i++)
            {
                string path = "C:\\PrintInvoice";
                path = string.Concat(path, "\\", invoiceRefGroup[i], ".pdf");
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
            }
            if (File.Exists(outputPdfPath))
            {
                File.Delete(outputPdfPath);
            }
        }

        private void simpleButton1_Click_1(object sender, EventArgs e)
        {
        }
    }
}
