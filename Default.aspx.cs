using ExcelLibrary.SpreadSheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Data.SqlClient;



namespace Marek_Ptak
{
    public partial class _Default : Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }
        protected void Button1_Click(object sender, EventArgs e)
        {

           
            string path = @"C:\Dostawy";
            
            // Sprawdzenie czy istnieje katalog docelowy
            if (Directory.Exists(path))
            {
                Console.WriteLine("Katalog już istnieje.");
            }
            else
            { Directory.CreateDirectory(path); }


            Random rnd = new Random();
            string[] gatunek =
            {
                "45", "X5CrNiMo17-12-2", "25CrMo4", "24CrMo5", "32CrMo12", "H23N13", "40CrMoV4-6"
            };
            double[] srednica =
            {
                10.5, 12.7, 6, 22.9, 24, 7.6, 40, 48.4, 54.3, 60, 78, 120
            };
            string[] wytop =
            {
                "143434","344333", "367856", "6867067","248454", "568355", "245744", "865654","044678", "794067", "469677"
            };
            string[] profil =
            {
                "pręt okrągły","pręt sześciokątny"
            };
            double[] masa =
            {
                506.5, 625, 76.6, 80, 190.3, 1050.7, 940.9, 260, 520, 250.4, 300
            };
            string[] dostawca =
            {
                "FPG","Arcellor", "HSW"
            };
            string[] dataDostawy =
            {
                "10.02.2020","15.02.2020", "06.02.2020", "30.01.2020"
            };
            string[] przyjal =
           {
                "Pracownik1","Pracownik2", "Pracownik3", "Pracownik4"
            };


            string file = @"C:\Dostawy\Dostawy materiałów.xlsx";

            Workbook workbook = new Workbook();
            Worksheet worksheet = new Worksheet("Ostatnia dostawa");

            //wypełnienie losowymi danymi wg podanego wzoru
            worksheet.Cells[0, 0] = new Cell("Id");
            worksheet.Cells[0, 1] = new Cell("Gatunek");
            worksheet.Cells[0, 2] = new Cell("Średnica [mm]");
            worksheet.Cells[0, 3] = new Cell("Wytop");
            worksheet.Cells[0, 4] = new Cell("Profil");
            worksheet.Cells[0, 5] = new Cell("Masa [kg]");
            worksheet.Cells[0, 6] = new Cell("Długość [m]");
            worksheet.Cells[0, 7] = new Cell("Dostawca");
            worksheet.Cells[0, 8] = new Cell("Data Dostawy");
            worksheet.Cells[0, 9] = new Cell("Przyjął");


            for (int i = 1; i <= 15; i++)
            {
                int IndexGatunek = rnd.Next(gatunek.Length);
                int IndexSrednica = rnd.Next(srednica.Length);
                int IndexWytop = rnd.Next(wytop.Length);
                int IndexProfil = rnd.Next(profil.Length);
                int IndexMasa = rnd.Next(masa.Length);
                int IndexDostawca = rnd.Next(dostawca.Length);
                int IndexDataDost = rnd.Next(dataDostawy.Length);
                int IndexPrzyjal = rnd.Next(przyjal.Length);

                double dlugosc1;
                string profil1 = profil[IndexProfil];
                double masa1 = masa[IndexMasa];
                double srednica1 = srednica[IndexSrednica];
                if (profil1 == "pręt okrągły")
                {
                     dlugosc1 = Math.Round((masa1 / 22.2) * ((60 * 60) / (srednica1 * srednica1)),1);
                }
                else
                {
                    dlugosc1 = 0;
                }




                worksheet.Cells[i, 0] = new Cell(i);
                worksheet.Cells[i, 1] = new Cell(gatunek[IndexGatunek]);
                worksheet.Cells[i, 2] = new Cell(srednica1);
                worksheet.Cells[i, 3] = new Cell(wytop[IndexWytop]);
                worksheet.Cells[i, 4] = new Cell(profil1);
                worksheet.Cells[i, 5] = new Cell(masa1);
                worksheet.Cells[i, 6] = new Cell(dlugosc1);
                worksheet.Cells[i, 7] = new Cell(dostawca[IndexDostawca]);
                //worksheet.Cells[i, 7] = new Cell(dataDostawy[IndexDataDost]);
                worksheet.Cells[i, 8] = new Cell(DateTime.Now, @"DD-MM-YYYY");
                worksheet.Cells[i, 9] = new Cell(przyjal[IndexPrzyjal]);

                             
            }
            
            
            //ustawienie szerokości trzech pierwszych kolumn
            worksheet.Cells.ColumnWidth[0] = 1500;
            worksheet.Cells.ColumnWidth[1] = 4500;
            worksheet.Cells.ColumnWidth[2] = 3000;
            worksheet.Cells.ColumnWidth[3] = 3000;
            worksheet.Cells.ColumnWidth[4] = 4500;
            worksheet.Cells.ColumnWidth[5] = 3000;
            worksheet.Cells.ColumnWidth[6] = 3000;
            worksheet.Cells.ColumnWidth[7] = 3000;
            worksheet.Cells.ColumnWidth[8] = 3000;
            worksheet.Cells.ColumnWidth[9] = 3000;




            workbook.Worksheets.Add(worksheet);

            //zapisanie pliku jeżeli nieistnieje plik o tej nazwie w tej lokalizacji
            //lub podanie komunikatu jeżeli istnieje już plik o tej nazwie
            if (File.Exists(file))
            {
                Label2.Text = "Plik Test.xlsx juz istnieje w tej lokalizacji.";
            }
            else
            { 
                workbook.Save(file);
                if (File.Exists(file))
                {
                    Label2.Text = "Plik Test.xlsx został zapisany";
                }
                else
                {
                    Label2.Text = "Jeżeli plik się nie zapisał proszę o skontaktowanie się z ...";
                }
            }
             
        }

        protected void Button2_Click(object sender, EventArgs e)
        {
            string file = @"C:\Dostawy\Dostawy materiałów.xlsx";
            Workbook book = Workbook.Load(file); 
            Worksheet sheet = book.Worksheets[0];
        }
    }
}