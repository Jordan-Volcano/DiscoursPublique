using System;
using System.Linq;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.Globalization;
using System.Threading;

namespace DiscoursPublique
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("fr-FR");

            var bosquejos = new Bosquejos();
            Workbook wb = Globals.ThisAddIn.GetActiveWorkbook();

            foreach (Excel.Worksheet sheet in wb.Worksheets)
            {
                if (sheet.Name == "Bosquejos") {
                    
                    Range rngBosquejos = sheet.get_Range("A5", "A236");
                    Range rngDateBosquejos = sheet.get_Range("C5", "C236");

                    System.Array valuesBosquejos= (System.Array)rngBosquejos.Cells.Value2;
                    System.Array valuesDateBosquejos = (System.Array)rngDateBosquejos.Cells.Value;

                    bosquejos = new Bosquejos(
                                                    Globals.ThisAddIn.ConvertToStringArray(valuesBosquejos),
                                                    Globals.ThisAddIn.ConvertToStringArray(valuesDateBosquejos)
                                                  );

                    // MessageBox.Show( bosquejos.numeros.ElementAt(i) ); // == bosquejo n° i
                    // MessageBox.Show( bosquejos.datesBosquejo.ElementAt(i) ); // date bosque n° i
                }

                else if(sheet.Name == "Hermanos") {
                   
                }
                else { 
                    Range rngdiscours = sheet.get_Range("D13", "D36");
                    Range rngdates = sheet.get_Range("A13", "A36");

                    System.Array valuesDiscours = (System.Array)rngdiscours.Cells.Value2;
                    System.Array valuesDates = (System.Array)rngdates.Cells.Value;

                    string[] discours = Globals.ThisAddIn.ConvertToStringArray(valuesDiscours);
                    string[] dates = Globals.ThisAddIn.ConvertToStringArray(valuesDates);

                    for (int i = 1; i < discours.Length; i++)
                    {
                        for (int j = 1; j < bosquejos.numeros.Length; j++) { 

                            if (discours[i] == bosquejos.numeros.ElementAt(j) )
                            {
                                                           
                                DateTime dateDiscours = DateTime.Parse(dates[i], new CultureInfo("fr-FR")) ;

                                if (bosquejos.datesBosquejo.ElementAt(j) == "")
                                {

                                    bosquejos.datesBosquejo[j] = "01/01/2015";
                                }

                                //TODO : Prendre en charge les valeur où il n'y a pas de date

                                DateTime dateBosquejo = DateTime.Parse(bosquejos.datesBosquejo.ElementAt(j), new CultureInfo("fr-FR"));
                               
                             

                               if( dateDiscours > dateBosquejo)
                                {


                                    MessageBox.Show("la date programmé "+ dateDiscours.ToShortDateString() +" est supérieur à la date inscrite "+ dateBosquejo.ToShortDateString() +" pour le discours n° "+ discours[i] );

                                    Worksheet sheet2 = wb.Sheets["Bosquejos"];

                                    Range newrange = sheet2.get_Range("C" + (j + 5));
                                    
                                    newrange.Cells.Value = dateDiscours.ToShortDateString();

                                }


                                
                                

                            }

                        }

                    }
                }
            }

            MessageBox.Show("Les dates des discours ont été mise à jour");

        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("fr-FR");

            var orateur = new Orateur();
            Workbook wb = Globals.ThisAddIn.GetActiveWorkbook();

            foreach (Excel.Worksheet sheet in wb.Worksheets)
            {
                if (sheet.Name == "Bosquejos") {
                    
                    Range rngOrateur = sheet.get_Range("C5", "P5");
                    Range rngDateOrateurs = sheet.get_Range("C4", "P4");

                    System.Array valuesOrateur= (System.Array)rngOrateur.Cells.Value;
                    System.Array valuesDateBosquejos = (System.Array)rngDateOrateurs.Cells.Value;

                    orateur = new Orateur(
                                                    Globals.ThisAddIn.ConvertToStringArray2(valuesOrateur),
                                                    Globals.ThisAddIn.ConvertToStringArray2(valuesDateBosquejos)
                                                  );

                    // MessageBox.Show( bosquejos.numeros.ElementAt(i) ); // == bosquejo n° i
                    // MessageBox.Show( bosquejos.datesBosquejo.ElementAt(i) ); // date bosque n° i
                }

                else if(sheet.Name == "Hermanos") {
                   
                }
                else { 
                    Range rngorateurs = sheet.get_Range("B13", "B36");
                    Range rngdates = sheet.get_Range("A13", "A36");

                    System.Array valuesOrateurs = (System.Array)rngorateurs.Cells.Value2;
                    System.Array valuesDates = (System.Array)rngdates.Cells.Value;

                    string[] listeorateurs = Globals.ThisAddIn.ConvertToStringArray(valuesOrateurs);
                    string[] dates = Globals.ThisAddIn.ConvertToStringArray(valuesDates);

                    for (int i = 1; i < listeorateurs.Length; i++)
                    {

                        if (dates[i] == "")
                        {
                            dates[i] = "01/01/2015";
                        }

                        for (int j = 1; j < orateur.Name.Length; j++) { 

                            if (listeorateurs[i] == orateur.Name.ElementAt(j) && listeorateurs[i] != "" )
                            {
                                                           
                                DateTime dateDiscours = DateTime.Parse(dates[i], new CultureInfo("fr-FR")) ;
                               
                                if (orateur.datesBosquejo[j] == "")
                                {

                                    orateur.datesBosquejo[j] = "01/01/2014";

                                   
                                }

                                //TODO : Prendre en charge les valeur où il n'y a pas de date
                                
                                DateTime dateBosquejo = DateTime.Parse(orateur.datesBosquejo.ElementAt(j), new CultureInfo("fr-FR"));

                               
                                if ( dateDiscours > dateBosquejo)
                                {
                                    
                                    string[] lettre = {"C","D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P"};
                                   

                                    Worksheet sheet2 = wb.Sheets["Bosquejos"];

                                    Range newrange = sheet2.get_Range(lettre[j] + 4);
                                    
                                    newrange.Cells.Value = dateDiscours.ToShortDateString();

                                }


                                
                                

                            }

                        }

                    }
                }
            }

            MessageBox.Show("Les dates pour les orateurs à été mis à jour");
        }
    }
}
