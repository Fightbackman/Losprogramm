/*
* Losprogramm
Copyright (C) 2013 Lukas Glitt and Olaf Matticzk

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program. If not, see <http://www.gnu.org/licenses/>
*/



using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing.Printing;
using System.Data.SqlClient;
using Excel = Microsoft.Office.Interop.Excel;
namespace Losprogramm
{
    public partial class Form1 : Form
    {
        DataTable winData = new DataTable();
        PrintDialog Dialog = new PrintDialog();
        String[] LoseTest;
        String[] GewinneTest;
        DataColumn C1 = new DataColumn("Losnummer", typeof(String));
        DataColumn C2 = new DataColumn("Gewinnummer", typeof(String));
        
        
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("Losprogramm\r Copyright (C) 2013 Lukas Glitt and Olaf Matticzk\r This program is free software: you can redistribute it and/or modify\r it under the terms of the GNU General Public License as published by\r the Free Software Foundation, either version 3 of the License, or\r any later version.\r This program is distributed in the hope that it will be useful,\r but WITHOUT ANY WARRANTY; without even the implied warranty of\r MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the\r GNU General Public License for more details.\r You should have received a copy of the GNU General Public License\r along with this program. If not, see <http://www.gnu.org/licenses/>\r Stimmen sie dieser Lizensierung zu?", "Lizensierung", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.No)
            {
                System.Windows.Forms.Application.Exit();
            }
            else if (dialogResult == DialogResult.Yes)
            {

            }
        }
        public void addColumn(DataColumn Clmn1, DataColumn Clmn2)
        {
            winData.Columns.AddRange(new DataColumn[] { Clmn1, Clmn2 });
        }
        private void button1_Click(object sender, EventArgs e)
        {
            winData.Clear();
            winData.Columns.Clear();
            addColumn(C1, C2);
            try
            {
                
                int xi = Int32.Parse(textBox1.Text);
                int yi = Int32.Parse(textBox2.Text);
                if((xi/100*yi)<1){
                    MessageBox.Show("Falsche Eingabe. Die Ergebnisse wären kleiner 1.");
                    textBox1.Text = "";
                    textBox2.Text = "";
                    return;
                }
                String[] Lose = new String[xi];
                String[] Gewinne = new String[xi * yi / 100];
                Lose = BubbleSort(zufallsvergabe(xi, yi));
                Gewinne = zufallsvergabe(xi * yi / 100, 100);
                LoseTest = Lose;
                GewinneTest = Gewinne;
                // textBox3.Text = "LosNmr" + "     " + "GewinnNmr" + '\r' + '\n' + AusgabeArrayTabelle(Lose, Gewinne);


                for (int i = 0; i < Lose.Length; i++)
                {


                    DataRow Reihe = winData.NewRow();
                    Reihe["Losnummer"] = Lose[i];
                    Reihe["Gewinnummer"] = Gewinne[i];
                    winData.Rows.Add(Reihe);

                }
                Table1.DataSource = winData;
                Table1.AutoSize = true;

            }
            catch(ArgumentNullException f)
            {
                MessageBox.Show("Falsche Eingabe");
            }
            catch(FormatException g){
                MessageBox.Show("Falsche Eingabe");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ExportToExcel();
        }

        private const int WS_SYSMENU = 0x80000;
        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.Style &= ~WS_SYSMENU;
                return cp;
            }
        }


        public String[] zufallsvergabe(int DrchlfAnzhl, int Winchance)
        {
            Boolean[] Durchlauf = new Boolean[DrchlfAnzhl + 1];
            String[] gewinnA = new String[DrchlfAnzhl * Winchance / 100];
            Boolean rdy = new Boolean();
            Random x = new Random();
            rdy = false;
            int wins = 0;
            int win = 0;
            while (rdy == false)
            {
                win = x.Next(1, DrchlfAnzhl + 1);
                if (Durchlauf[win] != true)
                {
                    Durchlauf[win] = true;
                    gewinnA[wins] = win.ToString();
                    wins = wins + 1;
                    if (wins == DrchlfAnzhl * Winchance / 100)
                        rdy = true;
                }


            }
            return gewinnA;


        }
        public String[] BubbleSort(String[] Array)
        {
            String[] Buffer = Array;
            for (int p = 0; p < Buffer.Length - 1; p++)
            {
                for (int z = 1; z <= (Buffer.Length - 1); z++)
                {
                    if (Int32.Parse(Buffer[z - 1]) > Int32.Parse(Buffer[z]))
                    {
                        String puffer = Buffer[z];
                        Buffer[z] = Buffer[z - 1];
                        Buffer[z - 1] = puffer;
                    }
                }
            }
            return Buffer;
        }

        public String AusgabeArray(Object[] Array)
        {
            String Asgb = "";
            for (int i = 0; i < Array.Length; i++)
            {
                Asgb = Asgb + '\r' + '\n' + Array[i];


            }
            return Asgb;
        }

        public String[] AusgabeArray2(Object[] Array, Object[] Array2)
        {
            String[] Asgb1 = new String[(Array2.Length / 28)+1];
            
            for (Double j = 0; j < (((Double)Array2.Length) / 28); j++)
            {
                String Asgb = "Losnummer:                                             Gewinnnummer:";
                if (j == ((int)(Array2.Length / 28)) && Array2.Length % 28 != 0)
                {
                    for (int k = 0; k < Array2.Length % 28; k++)
                    {
                        Asgb = Asgb + '\r' + '\n' + Array[k + ((int)(j * 28))] + "                       |                              " + Array2[k + ((int)(j * 28))] + '\r' + '\n' + "---------------------------------------------------------------------------------------------";
                    }
                }
                else
                {
                    for (int l = 0 + ((int)(j * 28)); l < (j + 1) * 28; l++ )
                    {
                        Asgb = Asgb + '\r' + '\n' + Array[l] + "                       |                              " + Array2[l] + '\r' + '\n' + "---------------------------------------------------------------------------------------------";
                    }

                }

                Asgb1[((int)j)] = Asgb;

                
                
             }
            return Asgb1;
        }

        public DataTable AusgabeArrayTabelle(Object[] Array, String Name1, Object[] Array2, String Name2)
        {

            DataTable Table = new DataTable();
            for (int i = 0; i < Array.Length; i++)
            {
                DataRow Reihe = Table.NewRow();
                Reihe[Name1] = Array[i];
                Reihe[Name2] = Array2[i];
                Table.Rows.Add(Reihe);

            }

            return Table;
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
        String[] BufferArray;
        private void button3_Click(object sender, EventArgs e)
        {
            x = -1;
            PrintDocument printDocument1 = new PrintDocument();
            List<string> Printers = new List<string>();
            foreach (string p in PrinterSettings.InstalledPrinters)
                Printers.Add(p);
            printDocument1.PrinterSettings.PrinterName = PrinterSettings.InstalledPrinters[0];
            printDocument1.PrintPage += new PrintPageEventHandler(PrintPage);
            Dialog.AllowSomePages = true;
            Dialog.ShowHelp = true;
            Dialog.Document = printDocument1;
            DialogResult result = Dialog.ShowDialog();
            BufferArray = AusgabeArray2(LoseTest, GewinneTest);
            if (result == DialogResult.OK)
            {
                printDocument1.Print();
            }
        }
        int x = -1;
        void PrintPage(object sender, PrintPageEventArgs e)
        {
            


            x++;
            e.Graphics.DrawString(BufferArray[x], new Font("Times New Roman", 12), new SolidBrush(Color.Black), new Point(45, 45));
            if (x + 1 < ((int)GewinneTest.Length / 28))
            {
                e.HasMorePages = true;
            }
            if (((Double)x) + 1 < ((Double)GewinneTest.Length) / 28)
            {
                e.HasMorePages = true;
            }
            
            
        }

        private void ExportToExcel()
        {
            //Excel Variablen deklarieren
            Excel.Application myExcelApplication;
            Excel.Workbook myExcelWorkbook;
            Excel.Worksheet myExcelWorksheet;
            myExcelApplication = null;


            try
            {
                //Excel Prozess starten
                myExcelApplication = new Excel.Application();
                myExcelApplication.Visible = false;
                myExcelApplication.ScreenUpdating = false;

                //Excel Datei anlegen
                var myCount = myExcelApplication.Workbooks.Count;
                myExcelWorkbook = (Excel.Workbook)(myExcelApplication.Workbooks.Add(System.Reflection.Missing.Value));
                myExcelWorksheet = (Excel.Worksheet)myExcelWorkbook.ActiveSheet;

                //Spalten Überschriften
                myExcelWorksheet.Cells[2, 2] = "Losnummer";
                myExcelWorksheet.Cells[2, 3] = "Gewinnummer";

                //Formatieren der Überschrift
                Excel.Range myRangeHeadline;
                myRangeHeadline = myExcelWorksheet.get_Range("B2", "C2");
                myRangeHeadline.Font.Bold = true;
                myRangeHeadline.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                //Daten eingeben
                int j = LoseTest.Length;
                int k = 0;
                for (int i = 3; i <j+3; i++)
                {
                    String l = LoseTest[k];
                    myExcelWorksheet.Cells[i, 2] = l;
                    k++;
                }

                j = GewinneTest.Length;
                k = 0;
                for (int i = 3; i < j + 3; i++)
                {
                    String l = GewinneTest[k];
                    myExcelWorksheet.Cells[i, 3] = l;
                    k++;
                }
            

                //Excel Datei abspeichern
                //wenn die Datei bereits existeirt kommt eine Fehelrmeldung
                String Dateiname = "C:\\Losungsergebnisse_vom_";
                String Datum = System.DateTime.Now.ToString();
                Datum = Datum.Replace(' ', '_');
                Datum = Datum.Replace(':', '_');
                Datum = Datum.Replace('.','_');
                Dateiname = Dateiname + Datum +".xls";
                myExcelWorkbook.Close(true, Dateiname, System.Reflection.Missing.Value);
                MessageBox.Show("Die Ergebnisse wurden erfolgreich unter: " + Dateiname + " gespeichert.");
            }

            catch (Exception ex)
            {
                String myErrorString = ex.Message;
                MessageBox.Show(myErrorString);
            }
            finally
            {
                //Excell beenden
                if (myExcelApplication != null)
                {
                    myExcelApplication.Quit();
                }
            }

        }

        private void button4_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("Wollen sie das Programm wirklich beenden? Falls sie nicht gespeichert haben, sind die Daten UNWIEDERBRINGLICH verloren!", "WARNUNG", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                System.Windows.Forms.Application.Exit();
            }
            else if (dialogResult == DialogResult.No)
            {

            }
        }


    }

}

    


