using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Plaksin006
{
    public class Experiment
    {
        internal const int P = 23;
        internal Random rndGen = new Random();

        public void Start(out int[,] BDs, out bool proves)
        {
            BDs = new int[2, P];
            proves = false;
            for (int i = 0; i < P; i++)
            {
                BDs[1, i] = rndGen.Next(1, 13);

                switch (BDs[1, i])
                {
                    case 1:
                    case 3:
                    case 5:
                    case 7:
                    case 8:
                    case 10:
                    case 12:
                        BDs[0, i] = rndGen.Next(1, 32);
                        break;
                    case 4:
                    case 6:
                    case 9:
                    case 11:
                        BDs[0, i] = rndGen.Next(1, 31);
                        break;
                    case 2:
                        BDs[0, i] = rndGen.Next(1, 29);
                        break;
                }

                for (int j = 0; j < i; j++)
                    if (BDs[1, i] == BDs[1, j])
                        if (BDs[0, i] == BDs[0, j])
                            proves = true;
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            var try1 = new Experiment();

            const int N = 50;
            int cell = 1;

            var excelApp = new Excel.Application();
            excelApp.Visible = true;
            excelApp.Workbooks.Add();
            Excel._Worksheet workSheet = (Excel.Worksheet) excelApp.ActiveSheet;
            workSheet.Cells[1, cell] = "Эксперимент №";
            workSheet.Cells[2, cell] = "Есть совпадения?";
            workSheet.Cells[4, cell] = "Дата Рождения";

            for (int i = 0; i < N; i++)
            {
                cell++;
                int[,] BD;
                bool proven;
                try1.Start(out BD, out proven);

                workSheet.Cells[1, cell] = i + 1;
                workSheet.Cells[2, cell] = proven;
                for (int j = 0; j < BD.Length/2; j++)
                    workSheet.Cells[j + 4, cell] = "'" + BD[0, j].ToString("00") + "." + BD[1, j].ToString("00");
            }
        }
    }
}
