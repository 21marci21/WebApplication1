using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using WinFormsApp1.Models;

namespace WinFormsApp1
{
    public partial class Form1 : Form
    {

        Excel.Application xlApp; // A Microsoft Excel alkalmaz�s
        Excel.Workbook xlWB;     // A l�trehozott munkaf�zet
        Excel.Worksheet xlSheet; // Munkalap a munkaf�zeten bel�l

        public Form1()
        {
            InitializeComponent();
        }


        void CreateExcel()
        {
            try
            {
                // Excel elind�t�sa �s az applik�ci� objektum bet�lt�se
                xlApp = new Excel.Application();

                // �j munkaf�zet
                xlWB = xlApp.Workbooks.Add(Missing.Value);

                // �j munkalap
                xlSheet = xlWB.ActiveSheet;

                // T�bla l�trehoz�sa
                CreateTable(); // Ennek meg�r�sa a k�vetkez� feladatr�szben k�vetkezik

                // Control �tad�sa a felhaszn�l�nak
                xlApp.Visible = true;
                xlApp.UserControl = true;
            }
            catch (Exception ex) // Hibakezel�s a be�p�tett hiba�zenettel
            {
                string errMsg = string.Format("Error: {0}\nLine: {1}", ex.Message, ex.Source);
                MessageBox.Show(errMsg, "Error");

                // Hiba eset�n az Excel applik�ci� bez�r�sa automatikusan
                xlWB.Close(false, Type.Missing, Type.Missing);
                xlApp.Quit();
                xlWB = null;
                xlApp = null;
            }
        }

        void CreateTable()
        {
            string[] fejl�cek = new string[]
            {
                "K�rd�s",
                "1. V�lasz",
                "2. V�lasz",
                "3. V�lasz",
                "Helyes v�lasz",
                "K�p"
            };

            for (int  i = 0;  i < fejl�cek.Length;  i++)
            {
                xlSheet.Cells[1, i + 1] = fejl�cek[i];
            }

            HajosContext context = new HajosContext();
            var mindenKerdes = context.Questions.ToList();

            object[,] adatT�mb = new object[mindenKerdes.Count(), fejl�cek.Count()];

            for (int i = 0; i < mindenKerdes.Count(); i++)
            {
                adatT�mb[i, 0] = mindenKerdes[i].Question1;
                adatT�mb[i, 1] = mindenKerdes[i].Answer1;
                adatT�mb[i, 2] = mindenKerdes[i].Answer2;
                adatT�mb[i, 3] = mindenKerdes[i].Answer3;
                adatT�mb[i, 4] = mindenKerdes[i].CorrectAnswer;
                adatT�mb[i, 5] = mindenKerdes[i].Image;
            }

            int soroksz�ma = adatT�mb.GetLength(0);
            int oszlopoksz�ma = adatT�mb.GetLength(1);

            Excel.Range adatrange = xlSheet.get_Range("A2", Type.Missing).get_Resize(soroksz�ma, oszlopoksz�ma);
            adatrange.Value2 = adatT�mb;
            adatrange.Columns.AutoFit();
            adatrange.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick);

            Excel.Range elsooszlop = xlSheet.get_Range("A2", Type.Missing).get_Resize(soroksz�ma, 1);
            elsooszlop.Font.Bold = true;
            elsooszlop.Interior.Color = Color.LightYellow;

            Excel.Range utolsooszlop = xlSheet.get_Range("E2", Type.Missing).get_Resize(soroksz�ma, 1);
            utolsooszlop.Interior.Color = Color.LightGreen;
            utolsooszlop.NumberFormat= "#.00";

            Excel.Range fejll�cRange = xlSheet.get_Range("A1", Type.Missing).get_Resize(1, 6);
            fejll�cRange.Font.Bold = true;
            fejll�cRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            fejll�cRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            fejll�cRange.EntireColumn.AutoFit();
            fejll�cRange.RowHeight = 40;
            fejll�cRange.Interior.Color = Color.Fuchsia;
            fejll�cRange.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick);
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