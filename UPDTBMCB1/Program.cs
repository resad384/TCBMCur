using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbobsCOM;
using System.Globalization;
using System.Xml;

namespace UPDTBMCB1
{
    internal class Program
    {
        private static string _usd;
        private static string _euro;
        private static string _gbp;
        private static string _rub;
        private static CultureInfo _cultureInfo = new CultureInfo("en-US", true);
        private const string _today = "http://www.tcmb.gov.tr/kurlar/today.xml";


        private static void Main(string[] args)
        {
            var oCompany = new Company();
            oCompany.Server = "cassa.inci.az:30015";
            Console.WriteLine("Server:" + oCompany.Server + "");
            oCompany.DbServerType = BoDataServerTypes.dst_HANADB;
            oCompany.CompanyDB = "CASSA_PRD_NEW";
            Console.WriteLine("DB:" + oCompany.CompanyDB + "");
            oCompany.Password = "1234";
            oCompany.UserName = "manager";
            oCompany.language = BoSuppLangs.ln_English;
            oCompany.DbUserName = "B1DBU";
            oCompany.DbPassword = "Adm1nC@ss@123";






            try
            {
                oCompany.Connect();
                Console.WriteLine("Connected to company: " + oCompany.Connected);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                //throw;
            }


            {
   
                if (oCompany.Connected)
                {
                    try
                    {
                        XmlDocument _xmlDoc = new XmlDocument();
                        SAPbobsCOM.SBObob oSBObob;
                        oSBObob = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);

                        _xmlDoc.Load(_today);

                        Console.WriteLine(DateTime.Today.AddDays(1));
                        _usd = _xmlDoc.SelectSingleNode("Tarih_Date/Currency[@Kod='USD']/ForexBuying").InnerXml;
                        Console.WriteLine(_usd);
                        _euro = _xmlDoc.SelectSingleNode("Tarih_Date/Currency[@Kod='EUR']/ForexBuying").InnerXml;
                        Console.WriteLine(_euro);
                        _gbp = _xmlDoc.SelectSingleNode("Tarih_Date/Currency[@Kod='GBP']/ForexBuying").InnerXml;
                        Console.WriteLine(_gbp);
                        _rub = _xmlDoc.SelectSingleNode("Tarih_Date/Currency[@Kod='RUB']/ForexBuying").InnerXml;
                        Console.WriteLine(_rub);

                        if (_usd != "")
                        {
                            oSBObob.SetCurrencyRate("USD", DateTime.Today.AddDays(1),
                            Convert.ToDouble(_usd, CultureInfo.CurrentUICulture),
                            true);
                        }

                        if (_euro != "")
                        {
                            oSBObob.SetCurrencyRate("EUR", DateTime.Today.AddDays(1),
                             Convert.ToDouble(_euro, CultureInfo.CurrentUICulture),
                             true);
                        }

                        if (_gbp != "")
                        {
                            oSBObob.SetCurrencyRate("GBP", DateTime.Today.AddDays(1),
                             Convert.ToDouble(_gbp, CultureInfo.CurrentUICulture),
                             true);
                        }

                        if (_rub != "")
                        {
                            oSBObob.SetCurrencyRate("RUB", DateTime.Today.AddDays(1),
                    Convert.ToDouble(_rub, CultureInfo.CurrentUICulture),
                    true);
                        }



                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex);
                        System.Threading.Thread.Sleep(5000);
                    }
                }

            }
        }
    }
}
