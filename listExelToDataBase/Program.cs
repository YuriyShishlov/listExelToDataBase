using listExelToDataBase.Data.DataModel;
using System;
using System.Data.SQLite;
using System.Linq;

namespace listExelToDataBase
{
    class Program
    {
        static void Main(string[] args)
        {
            //Строка подключения к SQLite:
            //string sqliteConnection = @"F:\Projects\UCIT\listExelToDataBase\listExelToDataBase\Files\lpuSQLite.sqlite";
            //using (var context = new SQLiteConnection(string.Format("Data Source={0}; Version=3;", sqliteConnection)))
            //{
            //    var lpuData = context.LpuDatas.Where(p => p.NameHead == "Сергей");
                
            //    Console.WriteLine("В базе {0} Сергеев", lpuData.Count());
            //}
            //using (LpuContext context = new LpuContext())
            //{
            //    var lpuData = context.LpuDatas.Where(p => p.NameHead == "Сергей");

            //    Console.WriteLine("В базе {0} Сергеев", lpuData.Count());
            //}

            Console.ReadLine();
        }
    }
}
//int i = context.LpuDatas.Count();

//var lpudata = context.LpuDatas.ToList();

//LpuData lpuData = context.LpuDatas.Find(9395);

//foreach (var item in lpuData)
//{
//    Console.WriteLine(item.NameHead);
//}
//Console.WriteLine(lpuData.SurnameHead);