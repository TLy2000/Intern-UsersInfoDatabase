using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UserDatabase
{
    class Program
    {
        static void Main(string[] args)
        {
            string fName, lName, prmpt;
            int cnt = 0;
            List<string> fNList = new List<string>();
            List<string> lNList = new List<string>();
            do
            {
                fName = Person.GetFirstName();
                lName = Person.GetLastName();
                prmpt = Person.GetPrompt();
                cnt++;
                fNList.Add(fName);
                lNList.Add(lName);
            } while (prmpt != "n" && prmpt != "N");
            string filePath = @"C:\temp\temptest23.xlsx";
            Excel excel = new Excel(filePath, 1);
            excel.WriteRange(cnt, fNList);
            excel.WriteRange(cnt, lNList);
            int col = excel.CountColumns(1);
            int row = excel.CountRows(1);
            excel.ReadRange(row, col);
            excel.Close();

            Console.ReadKey();
        }
    }
}

