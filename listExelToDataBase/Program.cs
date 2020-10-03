using listExelToDataBase.Data;
using listExelToDataBase.Data.DataModel;
using System;
using System.Data.OleDb;
using System.Linq;

namespace listExelToDataBase
{
    class Program
    {
        static void Main(string[] args)
        {
            ////SQL Server
            ////OleDbConnection connect = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0; Data Source=F:\Projects\UCIT\listExelToDataBase\listExelToDataBase\Files\listExel.xls; Extended Properties=""Excel 8.0""");
            //using (OleDbConnection connect = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0; Data Source=F:\Projects\UCIT\listExelToDataBase\listExelToDataBase\Files\listExel.xls; Extended Properties='Excel 8.0'"))
            //{
            //    connect.Open();

            //    OleDbCommand cmd = new OleDbCommand("SELECT * FROM [Лист1$]", connect);

            //    OleDbDataReader reader = cmd.ExecuteReader();

            //    while (reader.Read())
            //    {
            //        using (LpuContext context = new LpuContext())
            //        {
            //            LpuData lpuData = new LpuData();
            //            lpuData.NameSubjectRussianFederation = reader["NameSubjectRussianFederation"].ToString();
            //            lpuData.YearActivityCarriedOut = reader["YearActivityCarriedOut"].ToString();
            //            lpuData.MedicalOrganizationCode = reader["MedicalOrganizationCode"].ToString();
            //            lpuData.FullNameMedicalOrganization = reader["FullNameMedicalOrganization"].ToString();
            //            lpuData.ShortNameMedicalOrganization = reader["ShortNameMedicalOrganization"].ToString();
            //            lpuData.OrganizationalLegalForm = reader["OrganizationalLegalForm"].ToString();
            //            lpuData.Postcode = reader["Postcode"].ToString();
            //            lpuData.AddressMedicalOrganization = reader["AddressMedicalOrganization"].ToString();
            //            lpuData.SurnameHead = reader["SurnameHead"].ToString();
            //            lpuData.NameHead = reader["NameHead"].ToString();
            //            lpuData.PatronymicHead = reader["PatronymicHead"].ToString();
            //            lpuData.Phone = reader["Phone"].ToString();
            //            lpuData.Fax = reader["Fax"].ToString();
            //            lpuData.Email = reader["Email"].ToString();
            //            lpuData.LicenseNumber = reader["LicenseNumber"].ToString();
            //            lpuData.LicenseIssueDate = reader["LicenseIssueDate"].ToString();
            //            lpuData.LicenseExpirationDate = reader["LicenseExpirationDate"].ToString();
            //            lpuData.TypesMedicalCare = reader["TypesMedicalCare"].ToString();
            //            lpuData.DateInclusionInRegister = reader["DateInclusionInRegister"].ToString();

            //            context.LpuDatas.Add(lpuData);
            //            context.SaveChanges();
            //        }
            //    }
            //    connect.Close();
            //}

            //Console.WriteLine("Операция завершена");

            //SQLite
            //OleDbConnection connect = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0; Data Source=F:\Projects\UCIT\listExelToDataBase\listExelToDataBase\Files\listExel.xls; Extended Properties=""Excel 8.0""");
            //using (OleDbConnection connect = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0; Data Source=F:\Projects\UCIT\listExelToDataBase\listExelToDataBase\Files\listExel.xls; Extended Properties='Excel 8.0'"))
            //{
            //    connect.Open();

            //    OleDbCommand cmd = new OleDbCommand("SELECT * FROM [Лист1$]", connect);

            //    OleDbDataReader reader = cmd.ExecuteReader();

            //    while (reader.Read())
            //    {
            //        using (LpuContext context = new LpuContext())
            //        {
            //            LpuData lpuData = new LpuData();
            //            lpuData.NameSubjectRussianFederation = reader["NameSubjectRussianFederation"].ToString();
            //            lpuData.YearActivityCarriedOut = reader["YearActivityCarriedOut"].ToString();
            //            lpuData.MedicalOrganizationCode = reader["MedicalOrganizationCode"].ToString();
            //            lpuData.FullNameMedicalOrganization = reader["FullNameMedicalOrganization"].ToString();
            //            lpuData.ShortNameMedicalOrganization = reader["ShortNameMedicalOrganization"].ToString();
            //            lpuData.OrganizationalLegalForm = reader["OrganizationalLegalForm"].ToString();
            //            lpuData.Postcode = reader["Postcode"].ToString();
            //            lpuData.AddressMedicalOrganization = reader["AddressMedicalOrganization"].ToString();
            //            lpuData.SurnameHead = reader["SurnameHead"].ToString();
            //            lpuData.NameHead = reader["NameHead"].ToString();
            //            lpuData.PatronymicHead = reader["PatronymicHead"].ToString();
            //            lpuData.Phone = reader["Phone"].ToString();
            //            lpuData.Fax = reader["Fax"].ToString();
            //            lpuData.Email = reader["Email"].ToString();
            //            lpuData.LicenseNumber = reader["LicenseNumber"].ToString();
            //            lpuData.LicenseIssueDate = reader["LicenseIssueDate"].ToString();
            //            lpuData.LicenseExpirationDate = reader["LicenseExpirationDate"].ToString();
            //            lpuData.TypesMedicalCare = reader["TypesMedicalCare"].ToString();
            //            lpuData.DateInclusionInRegister = reader["DateInclusionInRegister"].ToString();

            //            context.LpuDatas.Add(lpuData);
            //            context.SaveChanges();
            //        }
            //    }
            //    connect.Close();
            //}

            //Console.WriteLine("Операция завершена");


            Console.ReadLine();
        }
    }
}
