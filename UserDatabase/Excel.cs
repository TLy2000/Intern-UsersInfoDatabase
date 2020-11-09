using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace UserDatabase
{
    class Excel
    {
        _Excel.Application xlApp = new _Excel.Application();
        Workbook wb;
        Worksheet ws;
        string path = "";
        char aChar = 'A';

        public Excel(string path, int Sheet)
        {
            // creates the excel object
            this.path = path;
            if (!File.Exists(path))
            {
                Console.WriteLine("Creating new database.");
                wb = xlApp.Workbooks.Add(Type.Missing);
            }
            else
            {
                Console.WriteLine("Adding to database.");
                wb = xlApp.Workbooks.Open(path);
            }
            ws = wb.Worksheets[Sheet];
        }
        public void ReadRange(int row, int col)
        {
            // reads multiple cells
            string cellsValue = "";
            // this loop increments each row
            for (int p = 1; p <= row; p++)
            {
                // this loop increments each column
                for (int q = 1; q <= col; q++)
                {
                    cellsValue = ws.Cells[p, q].Value2;
                    Console.Write(cellsValue);
                }
            }
        }

        public int CountColumns(int Sheet)
        {
            // this is to get the total nmber of columns used
            ws.Columns.ClearFormats();
            int iTotalColumns = ws.UsedRange.Columns.Count;
            return iTotalColumns;
        }

        public int CountRows(int Sheet)
        {
            // counts the rows used
            ws.Rows.ClearFormats();
            int iTotalRows = ws.UsedRange.Rows.Count;
            return iTotalRows;
        }
        public void WriteRange(int count, List<string> list)
        {
            // allows the user to write to multiple cells
            var wsItems = ws;
            ws.Cells[1, "A"] = "Unique ID";
            ws.Cells[1, "B"] = "First Name";
            ws.Cells[1, "C"] = "Last Name";
            // holds the cell to write to
            string cellName;
            Console.Write("Column: " + aChar);
            Console.WriteLine("Count: " + count);
            int counter = 2;
            cellName = aChar + counter.ToString();
            Console.WriteLine(cellName);
            if(ws.get_Range(cellName, cellName) == null)
            {
                Console.WriteLine("inside if. column: " + aChar + " Row: " + counter);
                foreach (var item in list)
                {
                    cellName = aChar + counter.ToString();
                    var range = ws.get_Range(cellName, cellName);
                    range.Value2 = item.ToString();
                    counter++;
                    Console.WriteLine("Column: " + aChar + "\tRow: " + counter + "\tValue: " + item);
                }
            }
            else
            {
                aChar++;
                Console.WriteLine("inside else. column: " + aChar + " Row: " + counter);
                foreach (var item in list)
                {
                    cellName = aChar + counter.ToString();
                    var range = ws.get_Range(cellName, cellName);
                    range.Value2 = item.ToString();
                    counter++;
                    Console.WriteLine("Column: " + aChar + "\tRow: " + counter + "\tValue: " + item);
                }
            }
        }

        public void Save()
        {
            // saves the file as is
            wb.Save();
        }

        public void SaveAs(string path)
        {
            // allows user to save the file to a selected directory or new file
            wb.SaveAs(path);
        }

        public void Close()
        {
            // closes the file
            wb.Close();
        }
    }
}
