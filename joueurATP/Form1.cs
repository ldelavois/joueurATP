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
using System.Configuration;

using System.Collections.Specialized;
using Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;

namespace joueurATP
{
    public partial class joueurATP : Form
    {
        //private string Excel07ConString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 8.0;HDR={1}'";
        private string filePath;
        private object sheetName;
        private DataTable dt;


        public joueurATP()
        {
            InitializeComponent();
        }

        private void joueurATP_Load(object sender, EventArgs e)
        {

        }

        //bouton Charger
        protected void btnCharger_Click(object sender, EventArgs e)
        {
            OpenFileDialog openfiledialog1 = new OpenFileDialog();
            openfiledialog1.ShowDialog();
            openfiledialog1.Filter = "allfiles|*.xlsx";
            txtExcelFile.Text = openfiledialog1.FileName;
            txtExcelFile.Refresh();
            string path = txtExcelFile.Text;


            filePath = openfiledialog1.FileName;
             //extension = Path.GetExtension(filePath);
            sheetName="Feuil1";

            //conStr = string.Format(Excel07ConString, filePath,0); //0 car pas de header

            ////Get the name of the First Sheet.
            //using (OleDbConnection con = new OleDbConnection(conStr))
            //    {
            //        using (OleDbCommand cmd = new OleDbCommand())
            //        {
            //            cmd.Connection = con;
            //            con.Open();
            //        DataTable dt = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            //        sheetName = dt.Rows[0]["Table_Name"].ToString();
            //        con.Close();
            //        }
            //    }

                //Read Data from the First Sheet.
            //    using (OleDbConnection con = new OleDbConnection(conStr))
            //    {
            //        using (OleDbCommand cmd = new OleDbCommand())
            //        {
            //            using (OleDbDataAdapter oda = new OleDbDataAdapter())
            //            {

            //                cmd.CommandText = "SELECT * From [" + sheetName + "]";
            //                cmd.Connection = con;
            //                con.Open();
            //                oda.SelectCommand = cmd;
            //                oda.Fill(dt);
            //                con.Close();



            //        }
            //    }
            //}

        }


        //bouton ajouter
        private void btnAjouter_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application xlApp;
            Workbook xlBook;
            Range xlRange;
            Worksheet xlSheet;

            try
            {
                xlApp = new Microsoft.Office.Interop.Excel.Application();
                xlBook = xlApp.Workbooks.Open(filePath);//
                xlSheet = xlBook.Worksheets[sheetName]; //doute
                xlRange = xlSheet.UsedRange;

                DataRow row = null;
                for (int i = 1; i <= xlRange.Rows.Count; i++)
                {
                    if (i != 1)
                        row = dt.NewRow();
                    for (int j = 1; j <= xlRange.Columns.Count; j++)
                    {
                        if (i == 1)
                        {
                            dt.Columns.Add();
                            dt.Columns.Add(xlRange.Cells[1, j].value);

                        }
                        else
                            row[j - 1] = xlRange.Cells[i, j].value;
                    }
                    if (row != null)
                        dt.Rows.Add(row);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                //xlBook.Close();
                //xlApp.Quit();
            }


            dt.Columns.AddRange(new DataColumn[5] {
        new DataColumn("jeu_id", typeof(int)),
        new DataColumn("naissance_date", typeof(DateTime)),
        new DataColumn("nom_joueur",typeof(string)),
        new DataColumn("is_grand_slam",typeof(bool)),
        new DataColumn("nb_victoire",typeof(int))
    });





            if (dt.Rows.Count > 0){
                string str = ConfigurationManager.ConnectionStrings["conString"].ConnectionString;
                using (SqlConnection con = new SqlConnection(str))
                {
                    using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(con))
                    {
                        sqlBulkCopy.DestinationTableName = "dbo.Jeu";
                        sqlBulkCopy.ColumnMappings.Add("jeu_id", "jeu_id");
                        sqlBulkCopy.ColumnMappings.Add("naissance_date", "naissance_date");
                        sqlBulkCopy.ColumnMappings.Add("nom_joueur", "nom_joueur");
                        sqlBulkCopy.ColumnMappings.Add("is_grand_slam", "is_grand_slam");
                        sqlBulkCopy.ColumnMappings.Add("nb_victoire", "nb_victoire");
                        con.Open();
                        sqlBulkCopy.WriteToServer(dt);
                        con.Close();
                    }
                }
            }
        }

    }
}
