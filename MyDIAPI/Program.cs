using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyDIAPI
{
    class Program
    {
        static void Main(string[] args)
        {
            SBO.Connection connection = new SBO.Connection();
            connection.Connect();

            long returnCode;
            string errMsg = string.Empty;

            try
            {
                SAPbobsCOM.BusinessPartners businessPartner = (SAPbobsCOM.BusinessPartners)connection.company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners);
                businessPartner.CardCode = "C06524";
                businessPartner.CardName = "Test Customer";
                businessPartner.CardType = SAPbobsCOM.BoCardTypes.cCustomer;

                businessPartner.CommissionGroupCode = 0;
                businessPartner.CommissionPercent = 15;
                businessPartner.CompanyPrivate = SAPbobsCOM.BoCardCompanyTypes.cPrivate;
                businessPartner.ContactPerson = "C1";
                businessPartner.Currency = "AUD";
                businessPartner.DiscountPercent = 15;
                businessPartner.VatLiable = SAPbobsCOM.BoVatStatus.vLiable;
                businessPartner.ShippingType = 3;

                businessPartner.ContactEmployees.Add();
                businessPartner.ContactEmployees.SetCurrentLine(0);
                businessPartner.ContactEmployees.Name = "C1";
                businessPartner.ContactEmployees.Address = "LN";
                businessPartner.ContactEmployees.E_Mail = "c1@abcd.com";
                businessPartner.ContactEmployees.Fax = "8433777778";
                businessPartner.ContactEmployees.MobilePhone = "8388888";
                businessPartner.ContactEmployees.Phone1 = "88880000";

                businessPartner.ContactEmployees.SetCurrentLine(1);
                businessPartner.ContactEmployees.Name = "C2";
                businessPartner.ContactEmployees.Address = "BJ";
                businessPartner.ContactEmployees.E_Mail = "c2@abcd.com";
                businessPartner.ContactEmployees.Fax = "84338";
                businessPartner.ContactEmployees.MobilePhone = "877388888";
                businessPartner.ContactEmployees.Phone1 = "8888300";

                returnCode = businessPartner.Add();
                if (returnCode == 0)
                {
                    connection.company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                }
                else
                {
                    if (connection.company.InTransaction)
                    {
                        connection.company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                    }

                    errMsg = connection.company.GetLastErrorDescription();
                }
            }
            catch (Exception ex)
            {

                Console.WriteLine(ex.Message);
            }

            Console.Read();
        }
    }
}
