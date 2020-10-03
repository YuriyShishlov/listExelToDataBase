using listExelToDataBase.Data;
using listExelToDataBase.Data.DataModel;
using System;
using System.Data.OleDb;
using System.Linq;
using System.Data.SQLite;

namespace listExelToDataBase
{
    class Program
    {
        static void Main(string[] args)
        {
            int numberInsert = 0;//для посчета добавленных строк

            //SQLite

            //Строка подключения к SQLite:
            string sqliteConnection = @"F:\Projects\UCIT\listExelToDataBase\listExelToDataBase\Files\lpuSQLite.sqlite";
            //OleDbConnection connect = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0; Data Source=F:\Projects\UCIT\listExelToDataBase\listExelToDataBase\Files\listExel.xls; Extended Properties=""Excel 8.0""");
            using (OleDbConnection connect = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0; Data Source=F:\Projects\UCIT\listExelToDataBase\listExelToDataBase\Files\listExel.xls; Extended Properties='Excel 8.0'"))
            {
                connect.Open();

                OleDbCommand cmd = new OleDbCommand("SELECT * FROM [Лист1$]", connect);

                OleDbDataReader reader = cmd.ExecuteReader();

                while (reader.Read())
                {
                    using (var con = new SQLiteConnection(string.Format("Data Source={0}; Version=3;", sqliteConnection)))
                    {
                        con.Open();

                        string nameSubjectRussianFederation = reader["NameSubjectRussianFederation"].ToString();
                        string yearActivityCarriedOut = reader["YearActivityCarriedOut"].ToString();
                        string medicalOrganizationCode = reader["MedicalOrganizationCode"].ToString();
                        string fullNameMedicalOrganization = reader["FullNameMedicalOrganization"].ToString();
                        string shortNameMedicalOrganization = reader["ShortNameMedicalOrganization"].ToString();
                        string organizationalLegalForm = reader["OrganizationalLegalForm"].ToString();
                        string postcode = reader["Postcode"].ToString();
                        string addressMedicalOrganization = reader["AddressMedicalOrganization"].ToString();
                        string surnameHead = reader["SurnameHead"].ToString();
                        string nameHead = reader["NameHead"].ToString();
                        string patronymicHead = reader["PatronymicHead"].ToString();
                        string phone = reader["Phone"].ToString();
                        string fax = reader["Fax"].ToString();
                        string email = reader["Email"].ToString();
                        string licenseNumber = reader["LicenseNumber"].ToString();
                        string licenseIssueDate = reader["LicenseIssueDate"].ToString();
                        string licenseExpirationDate = reader["LicenseExpirationDate"].ToString();
                        string typesMedicalCare = reader["TypesMedicalCare"].ToString();
                        string dateInclusionInRegister = reader["DateInclusionInRegister"].ToString();

                        string insertTableLpuDatas = string.Format("INSERT INTO LpuDatas(NameSubjectRussianFederation, YearActivityCarriedOut, MedicalOrganizationCode, FullNameMedicalOrganization, ShortNameMedicalOrganization, OrganizationalLegalForm, Postcode, AddressMedicalOrganization, SurnameHead, NameHead, PatronymicHead, Phone, Fax, Email, LicenseNumber, LicenseIssueDate, LicenseExpirationDate, TypesMedicalCare, DateInclusionInRegister) VALUES('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}','{15}','{16}','{17}','{18}')", nameSubjectRussianFederation, yearActivityCarriedOut, medicalOrganizationCode, fullNameMedicalOrganization, shortNameMedicalOrganization, organizationalLegalForm, postcode, addressMedicalOrganization, surnameHead, nameHead, patronymicHead, phone, fax, email, licenseNumber, licenseIssueDate, licenseExpirationDate, typesMedicalCare, dateInclusionInRegister);
                        string insertNewRow = string.Format("INSERT INTO LpuDatas (NameSubjectRussianFederation, YearActivityCarriedOut, MedicalOrganizationCode, FullNameMedicalOrganization, ShortNameMedicalOrganization, OrganizationalLegalForm, Postcode, AddressMedicalOrganization, SurnameHead, NameHead, PatronymicHead, Phone, Fax, Email, LicenseNumber, LicenseIssueDate, LicenseExpirationDate, TypesMedicalCare, DateInclusionInRegister) VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}','{15}','{16}','{17}','{18}')", nameSubjectRussianFederation, yearActivityCarriedOut, medicalOrganizationCode, fullNameMedicalOrganization, shortNameMedicalOrganization, organizationalLegalForm, postcode, addressMedicalOrganization, surnameHead, nameHead, patronymicHead, phone, fax, email, licenseNumber, licenseIssueDate, licenseExpirationDate, typesMedicalCare, dateInclusionInRegister);
                        var SqliteCmd = new SQLiteCommand(insertTableLpuDatas, con);
                        numberInsert += SqliteCmd.ExecuteNonQuery();//Execute the SqliteCommand
                        Console.WriteLine("Добавлено строк: {0}", numberInsert);
                    }
                }
                connect.Close();
            }
            
            //Очистка Таблицы в базе SQLite
            //using (var con = new SQLiteConnection(string.Format("Data Source={0}; Version=3;", sqliteConnection)))
            //{
            //    con.Open();
            //    string cleanSqlite = string.Format("DELETE FROM LpuDatas");
            //    SQLiteCommand cmd = new SQLiteCommand(cleanSqlite, con);
            //    int cleanNumber = cmd.ExecuteNonQuery();
            //    Console.WriteLine("Удалено {0}", cleanNumber);
            //}

            Console.WriteLine("Операция завершена");

            Console.ReadLine();
        }
    }
}
