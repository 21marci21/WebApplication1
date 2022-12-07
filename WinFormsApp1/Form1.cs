using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using WinFormsApp1.Models;

namespace WinFormsApp1
{
    public partial class Form1 : Form
    {

        Excel.Application xlApp; // A Microsoft Excel alkalmazás
        Excel.Workbook xlWB;     // A létrehozott munkafüzet
        Excel.Worksheet xlSheet; // Munkalap a munkafüzeten belül

        public Form1()
        {
            InitializeComponent();
        }


        void CreateExcel()
        {
            try
            {
                // Excel elindítása és az applikáció objektum betöltése
                xlApp = new Excel.Application();

                // Új munkafüzet
                xlWB = xlApp.Workbooks.Add(Missing.Value);

                // Új munkalap
                xlSheet = xlWB.ActiveSheet;

                // Tábla létrehozása
                CreateTable(); // Ennek megírása a következõ feladatrészben következik

                // Control átadása a felhasználónak
                xlApp.Visible = true;
                xlApp.UserControl = true;
            }
            catch (Exception ex) // Hibakezelés a beépített hibaüzenettel
            {
                string errMsg = string.Format("Error: {0}\nLine: {1}", ex.Message, ex.Source);
                MessageBox.Show(errMsg, "Error");

                // Hiba esetén az Excel applikáció bezárása automatikusan
                xlWB.Close(false, Type.Missing, Type.Missing);
                xlApp.Quit();
                xlWB = null;
                xlApp = null;
            }
        }

        void CreateTable()
        {
            string[] fejlécek = new string[]
            {
                "Kérdés",
                "1. Válasz",
                "2. Válasz",
                "3. Válasz",
                "Helyes válasz",
                "Kép"
            };

            for (int  i = 0;  i < fejlécek.Length;  i++)
            {
                xlSheet.Cells[1, i + 1] = fejlécek[i];
            }

            HajosContext context = new HajosContext();
            var mindenKerdes = context.Questions.ToList();

            object[,] adatTömb = new object[mindenKerdes.Count(), fejlécek.Count()];

            for (int i = 0; i < mindenKerdes.Count(); i++)
            {
                adatTömb[i, 0] = mindenKerdes[i].Question1;
                adatTömb[i, 1] = mindenKerdes[i].Answer1;
                adatTömb[i, 2] = mindenKerdes[i].Answer2;
                adatTömb[i, 3] = mindenKerdes[i].Answer3;
                adatTömb[i, 4] = mindenKerdes[i].CorrectAnswer;
                adatTömb[i, 5] = mindenKerdes[i].Image;
            }

            int sorokszáma = adatTömb.GetLength(0);
            int oszlopokszáma = adatTömb.GetLength(1);

            Excel.Range adatrange = xlSheet.get_Range("A2", Type.Missing).get_Resize(sorokszáma, oszlopokszáma);
            adatrange.Value2 = adatTömb;
            adatrange.Columns.AutoFit();
            adatrange.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick);

            Excel.Range elsooszlop = xlSheet.get_Range("A2", Type.Missing).get_Resize(sorokszáma, 1);
            elsooszlop.Font.Bold = true;
            elsooszlop.Interior.Color = Color.LightYellow;

            Excel.Range utolsooszlop = xlSheet.get_Range("E2", Type.Missing).get_Resize(sorokszáma, 1);
            utolsooszlop.Interior.Color = Color.LightGreen;
            utolsooszlop.NumberFormat= "#.00";

            Excel.Range fejllécRange = xlSheet.get_Range("A1", Type.Missing).get_Resize(1, 6);
            fejllécRange.Font.Bold = true;
            fejllécRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            fejllécRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            fejllécRange.EntireColumn.AutoFit();
            fejllécRange.RowHeight = 40;
            fejllécRange.Interior.Color = Color.Fuchsia;
            fejllécRange.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            CreateExcel(); 
            CreateTable();
        }

        private string GetCell(int x, int y)
        {
            string ExcelCoordinate = "";
            int dividend = y;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                ExcelCoordinate = Convert.ToChar(65 + modulo).ToString() + ExcelCoordinate;
                dividend = (int)((dividend - modulo) / 26);
            }
            ExcelCoordinate += x.ToString();

            return ExcelCoordinate;
        }
    }
}