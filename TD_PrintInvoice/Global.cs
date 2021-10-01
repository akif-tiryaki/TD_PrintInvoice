using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TD_PrintInvoice
{
    class Global
    {
        public string invoiceRef { get; set; }
        public string einvoice { get; set; }
        public string firmNr { get; set; }
        public string periodNr { get; set; }
        public string userName { get; set; }
        public string connectUsername { get; set; }
        public string connectPassword { get; set; }
        public string connectFirmNr { get; set; }
        public string DBConnection()
        {
            string conString = "";
            try
            {
                RegistryKey key;
                key = Registry.LocalMachine.OpenSubKey("Software\\Wow6432Node\\TD_GizIntegration");
                string server = key.GetValue("SERVER").ToString();
                string database = key.GetValue("DATABASE").ToString();
                string username = key.GetValue("USERNAME").ToString();
                string encodedText = key.GetValue("PASSWORD").ToString();
                var encodedTextBytes = Convert.FromBase64String(encodedText);
                string passwordText = Encoding.UTF8.GetString(encodedTextBytes);
                string password = passwordText;
                connectUsername = key.GetValue("CONNECT_USERNAME").ToString();
                string encodedTextConnect = key.GetValue("CONNECT_PASSWORD").ToString();
                var encodedTextBytesConnect = Convert.FromBase64String(encodedTextConnect);
                string passwordTextConnect = Encoding.UTF8.GetString(encodedTextBytesConnect);
                connectPassword = passwordTextConnect;
                connectFirmNr = key.GetValue("CONNECT_FIRMNR").ToString();
                conString = "Data Source=" + server + ";Initial Catalog=" + database + ";User ID=" + username + ";Password=" + password + ";Integrated Security=False";
            }
            catch (Exception ex)
            {

            }
            return conString;
        }
        public string[] LogoArg(string[] args_)
        {
            int i = 0;
            string[] Arr = new string[10];
            foreach (string item in args_)
            {
                Arr[i] = item;
                i++;
            }
            invoiceRef = args_[1];
            einvoice = args_[2];
            firmNr = args_[3];
            periodNr = args_[4];
            userName = args_[5];
            return args_;
        }
    }
}
