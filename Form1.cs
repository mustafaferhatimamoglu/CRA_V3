//using CrystalDecisions.CrystalReports.Engine;
//using CrystalDecisions.ReportAppServer.ReportDefModel;
using CrystalDecisions.CrystalReports.Engine;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CRA_V3
{
    public partial class Form1 : Form
    {
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        public Form1(string[] args)
        {
            InitializeComponent();

            if (args.Count() < 1)
            {
                MessageBox.Show("NO INPUT");
                //Application.Exit();
                //System.Windows.Forms.Application.Exit();
                //this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
                folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
                FolderBrowserDialog folderDlg = new FolderBrowserDialog();
                //DialogResult result = folderDlg.ShowDialog();
                DialogResult result = folderDlg.ShowDialog();
                if (result == DialogResult.OK)
                {
                    //textBox1.Text = folderDlg.SelectedPath;
                    args = new string[1] { folderDlg.SelectedPath };
                    Environment.SpecialFolder root = folderDlg.RootFolder;
                }
                //return;
            }
            //var parameters = File.ReadLines("C:\\AyruzMain_SHARED\\Debug\\net7.0-windows\\Data_FlowControl\\1^3^2023-07-21-23-44-53\\Parameters").ToArray();
            //MessageBox.Show(args[0]);
            var parameters = File.ReadLines(args[0] + "\\Parameters").ToArray();

            //crystalReportViewer1.
            ReportDocument myDataReport = new ReportDocument();
            myDataReport.Load(@"C:\AyruzSoftware\REPORTS\FlowControlReport.rpt");
            myDataReport.SetParameterValue("test", "t1");
            myDataReport.SetParameterValue("test2", "t2");

            //var asd = GetDataTableFromCsv("C:\\AyruzMain_SHARED\\Debug\\net7.0-windows\\Data_FlowControl\\1^3^2023-07-21-23-44-53\\MAIN.csv", true);
            var asd = GetDataTableFromCsv(args[0] + "\\MAIN.csv", true);
            asd.Columns[1].ColumnName = "Flow";
            asd.Columns[2].ColumnName = "Frequency";
            asd.Columns[3].ColumnName = "K_Factor";
            asd.Columns[4].ColumnName = "Temperature";

            DataTable asd2 = null;
            asd.DefaultView.Sort = "Flow" + " " + "DESC";
            asd2 = asd.DefaultView.ToTable();
            //asd.Rows[0][1] = "Flow";
            //asd.Rows[0][2] = "Frequency";
            //asd.Rows[0][3] = "K_Factor";
            //asd.Rows[0][4] = "Temperature";
            myDataReport.SetDataSource(asd2);
            //asd.Rows[0][1] = "Flow";

            //var parameters = File.ReadLines("c:\\file.txt").ToArray();

            myDataReport.SetParameterValue("PartNumber", parameters[0]);
            myDataReport.SetParameterValue("SerialNumber", parameters[1]);
            myDataReport.SetParameterValue("Size", parameters[2]);
            myDataReport.SetParameterValue("Fluid", parameters[3]);
            myDataReport.SetParameterValue("FluidTemp", parameters[4]);
            myDataReport.SetParameterValue("TestStartDate", parameters[5]);



            crystalReportViewer1.ReportSource = myDataReport;
            crystalReportViewer1.Refresh();
        }

        static DataTable GetDataTableFromCsv(string path, bool isFirstRowHeader)
        {
            string header = isFirstRowHeader ? "Yes" : "No";

            string pathOnly = Path.GetDirectoryName(path);
            string fileName = Path.GetFileName(path);

            string sql = @"SELECT * FROM [" + fileName + "]";

            using (OleDbConnection connection = new OleDbConnection(
                      @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + pathOnly +
                      ";Extended Properties=\"Text;HDR=" + header + "\""))
            using (OleDbCommand command = new OleDbCommand(sql, connection))
            using (OleDbDataAdapter adapter = new OleDbDataAdapter(command))
            {
                DataTable dataTable = new DataTable();
                dataTable.Locale = CultureInfo.CurrentCulture;
                adapter.Fill(dataTable);
                return dataTable;
            }
        }
    }
}
