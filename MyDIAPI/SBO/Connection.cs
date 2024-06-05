using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyDIAPI.SBO
{
    class Connection
    {
        public SAPbobsCOM.Company company;

        public SAPbobsCOM.BoDataServerTypes SQLVersion(string sqlVersion)
        {
            if (sqlVersion == "2016")
            {
                return SAPbobsCOM.BoDataServerTypes.dst_MSSQL2016;
            }
            else if (sqlVersion == "2017")
            {
                return SAPbobsCOM.BoDataServerTypes.dst_MSSQL2017;
            }
            if (sqlVersion == "2019")
            {
                return SAPbobsCOM.BoDataServerTypes.dst_MSSQL2019;
            }

            return SAPbobsCOM.BoDataServerTypes.dst_MSSQL2016;
        }

        public void Connect()
        {
            int lretcode = 0;
            string errmsg = string.Empty;

            if (company == null)
            {
                company = new SAPbobsCOM.Company();
            }

            if (company.Connected)
            {
                company.Disconnect();
            }

            try
            {
                company.Server = ConfigurationManager.AppSettings.GetValues("Server")[0];
                company.CompanyDB = ConfigurationManager.AppSettings.GetValues("CompanyDB")[0];
                company.UserName = ConfigurationManager.AppSettings.GetValues("UserName")[0];
                company.Password = ConfigurationManager.AppSettings.GetValues("Password")[0];
                company.DbUserName = ConfigurationManager.AppSettings.GetValues("DbUserName")[0];
                company.DbPassword = ConfigurationManager.AppSettings.GetValues("DBPassword")[0];
                company.DbServerType = SQLVersion(ConfigurationManager.AppSettings.GetValues("SQLVersion")[0]);
                company.language = SAPbobsCOM.BoSuppLangs.ln_English;

                lretcode = company.Connect();

                if (lretcode != 0)
                {
                    errmsg = company.GetLastErrorDescription();
                    Console.WriteLine(errmsg);
                }
                else
                {
                    Console.WriteLine("Connected");
                }

            }
            catch (Exception)
            {

            }
        }
    }
}
