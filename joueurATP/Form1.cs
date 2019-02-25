using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;
using System.Data.OleDb;
using Microsoft.Office.Interop.Excel;

namespace joueurATP
{
    public partial class joueurATP : Form
    {
        public joueurATP()
        {
            InitializeComponent();
        }


        private void btnCharger_Click(object sender, EventArgs e)
        {
            string ConnString = @"Server=SPC142130; Database=BIOPELAGOS; Trusted_Connection=yes;";

            OpenFileDialog openfiledialog1 = new OpenFileDialog();
            DialogResult res = openfiledialog1.ShowDialog();
            openfiledialog1.Filter = "allfiles|*.xlsx";
            txtExcelFile.Text = openfiledialog1.FileName;

            if (res == DialogResult.Cancel)
            {
                this.Close();
            }
            else
            {
                string path = txtExcelFile.Text;

                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

                Workbook xlBook = xlApp.Workbooks.Open(path); ;
                Worksheet xlSheet = xlBook.Worksheets["Feuil1"]; ;
                Range xlRange = xlSheet.UsedRange;
                StreamWriter log = new StreamWriter(@"C:\Users\ludovicd\Documents\biopelagos\log.txt");
                log.WriteLine("Liste des erreurs:");


                bool isError = false;

                System.Data.DataTable dt = new System.Data.DataTable();

                try
                {

                    DataRow row = null;
                    for (int i = 1; i <= xlRange.Rows.Count; i++)
                    {
                        if (i != 1) //si pas header
                            row = dt.NewRow(); //nouvelle colonne
                        for (int j = 1; j <= xlRange.Columns.Count; j++) //de la première à dernière colonne
                        {
                            if (i == 1) //si header

                            { dt.Columns.Add(xlSheet.Cells[1, j].value); } //creation d'une colonne

                            else //sinon si pas header alors
                            {
                                // ----- Gestion erreurs id ------
                                if (j == 1)
                                {
                                    if (xlSheet.Cells[i, 1].value == null) //si id null
                                    {
                                        isError = true;
                                        log.WriteLine("\nErreur ligne " + i + "\t  | id null");
                                    }
                                    else
                                    {
                                        if (xlSheet.Cells[i, 1].value < 1) //si id <1
                                        {
                                            isError = true;
                                            log.WriteLine("\nErreur ligne " + i + "\t  | id inférieur à 1 (valeur = " + xlSheet.Cells[i, 1].value + ")");
                                        }

                                        else
                                        {

                                            for (int k = 2; k <= xlRange.Rows.Count; k++)
                                                if (xlSheet.Cells[i, 1].value == xlSheet.Cells[k, 1].value && k != i)
                                                {
                                                    isError = true;
                                                    log.WriteLine("\nErreur ligne " + i + "\t  | id identique avec id ligne " + k);
                                                }

                                        } //end else id>=1

                                    } //end id not null

                                }

                                // ----- Gestion erreurs date ------
                                if (j == 2)
                                {

                                    DateTime dateNow = new DateTime();
                                    DateTime dateLimit = new DateTime(1970, 01, 01, 00, 00, 00, 00);
                                    DateTime? nullDateTime = null;

                                    if (xlSheet.Cells[i, 2].value == nullDateTime) //si date null
                                    {
                                        isError = true;
                                        log.WriteLine("\nErreur ligne " + i + "\t  | date null");
                                    }

                                    else
                                    {
                                        dateNow = DateTime.Now;
                                        int compareNow = DateTime.Compare(dateNow, xlSheet.Cells[i, 2].value);
                                        int compareLimit = DateTime.Compare(dateLimit, xlSheet.Cells[i, 2].value);

                                        if (compareNow == -1) //si date dans le futur
                                        {
                                            isError = true;
                                            log.WriteLine("\nErreur ligne " + i + "\t  | date dans le futur (valeur = " + xlSheet.Cells[i, 2].value + ")");
                                        }
                                        if (compareLimit == 1) //si date vieille
                                        {
                                            isError = true;
                                            log.WriteLine("\nErreur ligne " + i + "\t  | date trop ancienne (valeur = " + xlSheet.Cells[i, 2].value + ")");
                                        }
                                    } //end else not null

                                } //end if (j==2)

                                // ----- Gestion erreurs name ------
                                if (j == 3)
                                {
                                    if (xlSheet.Cells[i, 3].value == null) //si name null
                                    {
                                        isError = true;
                                        log.WriteLine("\nErreur ligne " + i + "\t  | name null");
                                    }
                                    else
                                    {
                                        //name en majuscule si pas déja en majuscule
                                        if (xlSheet.Cells[i, 3].value != xlSheet.Cells[i, 3].value.ToUpper())
                                        {
                                            xlSheet.Cells[i, 3].value = xlSheet.Cells[i, 3].value.ToUpper();
                                        }

                                        bool flagName = false;

                                        for (int n = 2; n <= xlRange.Rows.Count; n++)
                                        {
                                            if (xlSheet.Cells[i, 3].value == xlSheet.Cells[n, 3].value && n > i)
                                            //si names identiques (sans redondance)
                                            {
                                                flagName = true;
                                                isError = true;
                                                log.WriteLine("\nErreur ligne " + i + "\t  | name identique avec name ligne " + n);
                                            }
                                        }

                                        if (flagName == false)
                                        {


                                            using (SqlConnection connection = new SqlConnection(ConnString))
                                            {
                                            connection.Open();
                                            SqlCommand cmdName = new SqlCommand("SELECT count(*) as cnt FROM [BIOPELAGOS].[dbo].[Jeu] WHERE[nom_joueur] ='" + xlSheet.Cells[i, 3].value + "';", connection);
                                            SqlDataReader reader = cmdName.ExecuteReader();

                                            while (reader.Read())
                                            {
                                                if (reader["cnt"].ToString() != "0")
                                                {
                                                    isError = true;
                                                    log.WriteLine("\nErreur ligne " + i + "\t  | name déjà dans la table SQL");
                                                }

                                            }
                                            reader.Close();
                                            }
                                        } //end if (flagName == false)

                                    } //end if not null

                                } //end if(j==3)

                                // ----- Gestion erreurs is_slam ------
                                if (j == 4)
                                {
                                    if (xlSheet.Cells[i, 4].value == null) //si is_slam null
                                    {
                                        isError = true;
                                        log.WriteLine("\nErreur ligne " + i + "\t  | is_slam null");
                                    }
                                }

                                // ----- Gestion erreurs nb_slam ------
                                if (j == 5)
                                {
                                    if (xlSheet.Cells[i, 5].value == null) //si nb_slam null
                                    {
                                        isError = true;
                                        log.WriteLine("\nErreur ligne " + i + "\t  | nb_slam null");
                                    }
                                    else
                                    {
                                        if (xlSheet.Cells[i, 5].value < 0) //si nb_slam negatif
                                        {
                                            isError = true;
                                            log.WriteLine("\nErreur ligne " + i + "\t  | nb_slam négatif (valeur: " + xlSheet.Cells[i, 5].value + ")");
                                        }
                                        if (xlSheet.Cells[i, 5].value > 40) //si nb_slam extrême
                                        {
                                            isError = true;
                                            log.WriteLine("\nErreur ligne " + i + "\t  | nb_slam extrême (valeur:" + xlSheet.Cells[i, 5].value + ")");
                                        }
                                        if ((xlSheet.Cells[i, 5].value > 0 && xlSheet.Cells[i, 4].value == false) || (xlSheet.Cells[i, 5].value == 0 && xlSheet.Cells[i, 4].value == true))
                                        { //si incoherence entre is_slam et nb_slam
                                            isError = true;
                                            log.WriteLine("\nErreur ligne " + i + "\t  | incohérence [is_slam,nb_slam] (" + xlSheet.Cells[i, 4].value + "," + xlSheet.Cells[i, 5].value + ")");
                                        }
                                    }
                                }

                                row[j - 1] = xlSheet.Cells[i, j].value; //ajout de la valeur dans la case
                            }
                        }

                        if (row != null)
                            dt.Rows.Add(row);
                    } //end boucle i

                } //end try

                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
                finally
                {
                    xlBook.Close();
                    xlApp.Quit();
                }




                if (isError == false)
                {
                    log.WriteLine("Aucune erreur");
                    log.Close();
                    using (SqlBulkCopy bulkCopy = new SqlBulkCopy(ConnString))
                    {
                        bulkCopy.DestinationTableName = "dbo.Jeu";
                        bulkCopy.ColumnMappings.Add("id", "jeu_id");
                        bulkCopy.ColumnMappings.Add("date", "naissance_date");
                        bulkCopy.ColumnMappings.Add("name", "nom_joueur");
                        bulkCopy.ColumnMappings.Add("is_slam", "is_grand_slam");
                        bulkCopy.ColumnMappings.Add("nb_slam", "nb_victoire");
                        bulkCopy.WriteToServer(dt);

                    } // end using

                    //ajouter fenêtre "fichier correct importé"
                    MessageBox.Show("Pas d'erreur, importation réussie", "Log joueurATP");
                    this.Close();

                } //end if (isError == false)

                else
                {
                    //ajouter fenêtre des logs
                    log.Close();
                    string stringlog = File.ReadAllText(@"C:\Users\ludovicd\Documents\biopelagos\log.txt");
                    MessageBox.Show(stringlog, "Log joueurATP");

                    this.Close();
                }

            } //end else

        } //end btnCharger_Click

    } // end public partial class joueurATP : Form

} //end namespace joueur ATP
