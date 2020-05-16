using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Globalization;
using Gizmox.Controls;
using JDataEngine;
using JurisAuthenticator;
using JurisUtilityBase.Properties;
using System.Data.OleDb;

namespace JurisUtilityBase
{
    public partial class UtilityBaseMain : Form
    {
        #region Private  members

        private JurisUtility _jurisUtility;

        #endregion

        #region Public properties

        public string CompanyCode { get; set; }

        public string JurisDbName { get; set; }

        public string JBillsDbName { get; set; }

        public int FldClient { get; set; }

        public int FldMatter { get; set; }

        #endregion

        #region Constructor

        public UtilityBaseMain()
        {
            InitializeComponent();
            _jurisUtility = new JurisUtility();
        }

        #endregion

        #region Public methods

        public void LoadCompanies()
        {
            var companies = _jurisUtility.Companies.Cast<object>().Cast<Instance>().ToList();
//            listBoxCompanies.SelectedIndexChanged -= listBoxCompanies_SelectedIndexChanged;
            listBoxCompanies.ValueMember = "Code";
            listBoxCompanies.DisplayMember = "Key";
            listBoxCompanies.DataSource = companies;
//            listBoxCompanies.SelectedIndexChanged += listBoxCompanies_SelectedIndexChanged;
            var defaultCompany = companies.FirstOrDefault(c => c.Default == Instance.JurisDefaultCompany.jdcJuris);
            if (companies.Count > 0)
            {
                listBoxCompanies.SelectedItem = defaultCompany ?? companies[0];
            }
        }

        #endregion

        #region MainForm events

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        private void listBoxCompanies_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_jurisUtility.DbOpen)
            {
                _jurisUtility.CloseDatabase();
            }
            CompanyCode = "Company" + listBoxCompanies.SelectedValue;
            _jurisUtility.SetInstance(CompanyCode);
            JurisDbName = _jurisUtility.Company.DatabaseName;
            JBillsDbName = "JBills" + _jurisUtility.Company.Code;
            _jurisUtility.OpenDatabase();
            if (_jurisUtility.DbOpen)
            {
                cbClient.SelectedIndex = -1;
                string CliIndex;
                cbClient.ClearItems();
                string SQLCli = "select Client from (select '* All' as Client union all select dbo.jfn_formatclientcode(clicode) + '   ' +  clireportingname as Client from Client where clisysnbr in (select matclinbr from matter where matsysnbr in (select splitfrommat from splitbill))) CLI order by Client";
                DataSet myRSCli = _jurisUtility.RecordsetFromSQL(SQLCli);

                if (myRSCli.Tables[0].Rows.Count == 0)
                    cbClient.SelectedIndex = 0;
                else
                {
                    foreach (DataTable table in myRSCli.Tables)
                    {

                        foreach (DataRow dr in table.Rows)
                        {
                            CliIndex = dr["Client"].ToString();
                            cbClient.Items.Add(CliIndex);
                        }
                    }
                }
            }

        }



        #endregion

        #region Private methods

        private void DoDaFix()
        {
            // Enter your SQL code here
            // To run a T-SQL statement with no results, int RecordsAffected = _jurisUtility.ExecuteNonQueryCommand(0, SQL);
            // To get an ADODB.Recordset, ADODB.Recordset myRS = _jurisUtility.RecordsetFromSQL(SQL);


            DialogResult dt = MessageBox.Show("Splits for selected client/matters will be removed.  Do you wish to continue?",
                      "Split Removal Confirmation", MessageBoxButtons.YesNo);

            if (dt == DialogResult.Yes)
            {
                string singleClient = this.cbClient.GetItemText(this.cbClient.SelectedItem).Split(' ')[0];
                string singleMatter = this.cbMatter.GetItemText(this.cbMatter.SelectedItem).Split(' ')[0];

                string s2 = "select dbo.jfn_formatclientcode(clicode) as Client, clireportingname as ClientName, dbo.jfn_formatmattercode(matcode) as Matter, Matreportingname as MatterName, pbmprebill as Prebill from prebillmatter inner join matter on pbmmatter=matsysnbr inner join client on matclinbr=clisysnbr where (clicode like '%' + '" + singleClient + "' or '" + singleClient + "'='*') and " +
                    "  (matcode like '%' + '" + singleMatter + "' or '" + singleMatter + "'='*') ";
                DataSet d2 = _jurisUtility.RecordsetFromSQL(s2);
                if (d2.Tables[0].Rows.Count == 0)
                {
                    DataTable dvgSource = (DataTable)dataGridView1.DataSource;
                    int dvgRow = dvgSource.Rows.Count;
                    int i = 1;
                    foreach (DataRow dg in dvgSource.Rows)

                    {
                        string frommat = dg["SplitFromMat"].ToString();
                        string tomat = dg["SplitToMat"].ToString();
                        string CCode = dg["Client"].ToString();
                        string MCode = dg["Matter"].ToString();
                        Cursor.Current = Cursors.WaitCursor;
                        string dsql = "Delete from splitbill where splitfrommat=cast(" + frommat.ToString() + " as int)";
                        _jurisUtility.ExecuteNonQuery(0, dsql);

                        string ssql = "update matter Set matsplitmethod=0 where matsysnbr=cast(" + frommat.ToString() + " as int)";
                        _jurisUtility.ExecuteNonQuery(0, ssql);

                        toolStripStatusLabel.Text = "Splits for Client/Matter(s) " + CCode + "/" + MCode + " removed.";
                        statusStrip.Refresh();
                        UpdateStatus("Split Removal Complete", i, dvgRow);
                        i = i + 1;
                    }
                    Cursor.Current = Cursors.Default;
                    toolStripStatusLabel.Text = "Splits for Client/Matter(s) " + singleClient + "/" + singleMatter + " removed.";
                    statusStrip.Refresh();
                    UpdateStatus("Split Removal Complete", 1, 1);
                    WriteLog("Split Removal Client/Matter(s) " + singleClient + "/" + singleMatter + " " + DateTime.Now.ToShortDateString());
                    Application.DoEvents();

                    MessageBox.Show("Splits removed  for Client/Matter(s) " + singleClient + "/" + singleMatter);

                    cbClient.SelectedIndex = -1;
                    cbMatter.SelectedIndex = -1;
                    dataGridView1.DataSource = null;
                    dataGridView1.Rows.Clear();

                    cbClient.SelectedIndex = -1;
                    string CliIndex;
                    cbClient.ClearItems();
                    string SQLCli = "select Client from (select '* All' as Client union all select dbo.jfn_formatclientcode(clicode) + '   ' +  clireportingname as Client from Client where clisysnbr in (select matclinbr from matter where matsysnbr in (select splitfrommat from splitbill))) CLI order by Client";
                    DataSet myRSCli = _jurisUtility.RecordsetFromSQL(SQLCli);

                    if (myRSCli.Tables[0].Rows.Count == 0)
                        cbClient.SelectedIndex = 0;
                    else
                    {
                        foreach (DataTable table in myRSCli.Tables)
                        {

                            foreach (DataRow dr in table.Rows)
                            {
                                CliIndex = dr["Client"].ToString();
                                cbClient.Items.Add(CliIndex);
                            }
                        }
                    }
                }
                else
                {
                    DialogResult dt3 = MessageBox.Show("Open prebills exist for one or more selected client/matters.  These must be deleted before proceeding.  Click yes to view a list of open prebills.",
                   "Open Prebill Alert", MessageBoxButtons.YesNo);

                    if (dt3 == DialogResult.Yes)
                    { //generates output of the report for before and after the change will be made to client


                        ReportDisplay rpds = new ReportDisplay(d2);
                        rpds.Show();
                    }
                    else
                    {
                        Cursor.Current = Cursors.Default;
                        toolStripStatusLabel.Text = "Process Cancelled";
                        statusStrip.Refresh();
                        UpdateStatus("Process Cancelled", 0, 0);

                        Application.DoEvents();
                    }


                }
            }
            else

            {
                Cursor.Current = Cursors.Default;
                toolStripStatusLabel.Text = "Process Cancelled";
                statusStrip.Refresh();
                UpdateStatus("Process Cancelled", 0, 0);

                Application.DoEvents();
            }

            
            

        }
        private bool VerifyFirmName()
        {
            //    Dim SQL     As String
            //    Dim rsDB    As ADODB.Recordset
            //
            //    SQL = "SELECT CASE WHEN SpTxtValue LIKE '%firm name%' THEN 'Y' ELSE 'N' END AS Firm FROM SysParam WHERE SpName = 'FirmName'"
            //    Cmd.CommandText = SQL
            //    Set rsDB = Cmd.Execute
            //
            //    If rsDB!Firm = "Y" Then
            return true;
            //    Else
            //        VerifyFirmName = False
            //    End If

        }

        private bool FieldExistsInRS(DataSet ds, string fieldName)
        {

            foreach (DataColumn column in ds.Tables[0].Columns)
            {
                if (column.ColumnName.Equals(fieldName, StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            return false;
        }


        private static bool IsDate(String date)
        {
            try
            {
                DateTime dt = DateTime.Parse(date);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool IsNumeric(object Expression)
        {
            double retNum;

            bool isNum = Double.TryParse(Convert.ToString(Expression), System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out retNum);
            return isNum; 
        }

        private void WriteLog(string comment)
        {
            var sql =
                string.Format("Insert Into UtilityLog(ULTimeStamp,ULWkStaUser,ULComment) Values('{0}','{1}', '{2}')",
                    DateTime.Now, GetComputerAndUser(), comment);
            _jurisUtility.ExecuteNonQueryCommand(0, sql);
        }

        private string GetComputerAndUser()
        {
            var computerName = Environment.MachineName;
            var windowsIdentity = System.Security.Principal.WindowsIdentity.GetCurrent();
            var userName = (windowsIdentity != null) ? windowsIdentity.Name : "Unknown";
            return computerName + "/" + userName;
        }

        /// <summary>
        /// Update status bar (text to display and step number of total completed)
        /// </summary>
        /// <param name="status">status text to display</param>
        /// <param name="step">steps completed</param>
        /// <param name="steps">total steps to be done</param>
        private void UpdateStatus(string status, long step, long steps)
        {
            labelCurrentStatus.Text = status;

            if (steps == 0)
            {
                progressBar.Value = 0;
                labelPercentComplete.Text = string.Empty;
            }
            else
            {
                double pctLong = Math.Round(((double)step/steps)*100.0);
                int percentage = (int)Math.Round(pctLong, 0);
                if ((percentage < 0) || (percentage > 100))
                {
                    progressBar.Value = 0;
                    labelPercentComplete.Text = string.Empty;
                }
                else
                {
                    progressBar.Value = percentage;
                    labelPercentComplete.Text = string.Format("{0} percent complete", percentage);
                }
            }
        }

        private void DeleteLog()
        {
            string AppDir = Path.GetDirectoryName(Application.ExecutablePath);
            string filePathName = Path.Combine(AppDir, "VoucherImportLog.txt");
            if (File.Exists(filePathName + ".ark5"))
            {
                File.Delete(filePathName + ".ark5");
            }
            if (File.Exists(filePathName + ".ark4"))
            {
                File.Copy(filePathName + ".ark4", filePathName + ".ark5");
                File.Delete(filePathName + ".ark4");
            }
            if (File.Exists(filePathName + ".ark3"))
            {
                File.Copy(filePathName + ".ark3", filePathName + ".ark4");
                File.Delete(filePathName + ".ark3");
            }
            if (File.Exists(filePathName + ".ark2"))
            {
                File.Copy(filePathName + ".ark2", filePathName + ".ark3");
                File.Delete(filePathName + ".ark2");
            }
            if (File.Exists(filePathName + ".ark1"))
            {
                File.Copy(filePathName + ".ark1", filePathName + ".ark2");
                File.Delete(filePathName + ".ark1");
            }
            if (File.Exists(filePathName ))
            {
                File.Copy(filePathName, filePathName + ".ark1");
                File.Delete(filePathName);
            }

        }

            

        private void LogFile(string LogLine)
        {
            string AppDir = Path.GetDirectoryName(Application.ExecutablePath);
            string filePathName = Path.Combine(AppDir, "VoucherImportLog.txt");
            using (StreamWriter sw = File.AppendText(filePathName))
            {
                sw.WriteLine(LogLine);
            }	
        }
        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            DoDaFix();
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void cbClient_SelectedIndexChanged(object sender, EventArgs e)
        {

            string singleClient = this.cbClient.GetItemText(this.cbClient.SelectedItem).Split(' ')[0];
            cbMatter.SelectedIndex = -1;
            string MatIndex;
            cbMatter.ClearItems();
        
            string SQLMat = "select matter from (select '* All' as Matter union all " + 
            " select dbo.jfn_formatmattercode(matcode) + '   ' +  matreportingname as Matter from Matter inner join client on matclinbr=clisysnbr  where matsysnbr in (select splitfrommat from splitbill)  and dbo.jfn_formatclientcode(clicode)='" + singleClient.ToString() + "') MAT  order by matter";
            DataSet myRSMat = _jurisUtility.RecordsetFromSQL(SQLMat);


            if (myRSMat.Tables[0].Rows.Count == 0)
                cbMatter.SelectedIndex = 0;
            else
            {
                foreach (DataTable table in myRSMat.Tables)
                {

                    foreach (DataRow dr in table.Rows)
                    {
                        MatIndex = dr["Matter"].ToString();
                        cbMatter.Items.Add(MatIndex);
                    }
                }
            }

            cbMatter.SelectedIndex = 0;
        }

        private void cbMatter_SelectedIndexChanged(object sender, EventArgs e)
        {
            string singleClient = this.cbClient.GetItemText(this.cbClient.SelectedItem).Split(' ')[0];
            string singleMatter = this.cbMatter.GetItemText(this.cbMatter.SelectedItem).Split(' ')[0];

            if (singleClient == "*" && singleMatter == "*")
            {
                string getCliMat = "select dbo.jfn_formatclientcode(Client.clicode) as Client, Client.Clireportingname as ClientName, dbo.jfn_formatmattercode(Matter.matcode) as Matter, matter.matreportingname as MatterName, dbo.jfn_formatclientcode(SC.clicode) as SplitClient " +
                    ", SC.clireportingname as SplitClientName, dbo.jfn_Formatmattercode(SM.matcode) as SplitMatter, SM.matreportingname as SplitMatterName, cast(cast(splittopcnt as decimal(5,2)) as varchar(10)) + '%' as SplitPercent, SplitFromMat, SplitToMat  " +
                    " from splitbill inner join matter on splitfrommat=matsysnbr  inner join client on matclinbr=clisysnbr    " +
                    " inner join matter SM on splittomat=SM.matsysnbr inner join client SC on SC.clisysnbr=SM.matclinbr  ";

                DataSet myCliMat = _jurisUtility.RecordsetFromSQL(getCliMat);

                dataGridView1.DataSource = myCliMat.Tables[0];
            }
            else
                       if (singleClient != "*" && singleMatter == "*")
            {
                string getCliMat = "select dbo.jfn_formatclientcode(Client.clicode) as Client, Client.Clireportingname as ClientName, dbo.jfn_formatmattercode(Matter.matcode) as Matter, matter.matreportingname as MatterName, dbo.jfn_formatclientcode(SC.clicode) as SplitClient " +
                    ", SC.clireportingname as SplitClientName, dbo.jfn_Formatmattercode(SM.matcode) as SplitMatter, SM.matreportingname as SplitMatterName, cast(cast(splittopcnt as decimal(5,2)) as varchar(10)) + '%' as SplitPercent, SplitFromMat, SplitToMat  " +
                    " from splitbill inner join matter on splitfrommat=matsysnbr  inner join client on matclinbr=clisysnbr    " +
                    " inner join matter SM on splittomat=SM.matsysnbr inner join client SC on SC.clisysnbr=SM.matclinbr where  dbo.jfn_formatclientcode(client.clicode) in ('" + singleClient.ToString() + "')   ";

                DataSet myCliMat = _jurisUtility.RecordsetFromSQL(getCliMat);

                dataGridView1.DataSource = myCliMat.Tables[0];
            }
            else
             
            {
                string getCliMat = "select dbo.jfn_formatclientcode(Client.clicode) as Client, Client.Clireportingname as ClientName, dbo.jfn_formatmattercode(Matter.matcode) as Matter, matter.matreportingname as MatterName, dbo.jfn_formatclientcode(SC.clicode) as SplitClient " +
                    ", SC.clireportingname as SplitClientName, dbo.jfn_Formatmattercode(SM.matcode) as SplitMatter, SM.matreportingname as SplitMatterName, cast(cast(splittopcnt as decimal(5,2)) as varchar(10)) + '%' as SplitPercent, SplitFromMat, SplitToMat  " +
                    " from splitbill inner join matter on splitfrommat=matsysnbr  inner join client on matclinbr=clisysnbr    " +
                    " inner join matter SM on splittomat=SM.matsysnbr inner join client SC on SC.clisysnbr=SM.matclinbr where  dbo.jfn_formatclientcode(client.clicode) in ('" + singleClient.ToString() + "') and dbo.jfn_formatmattercode(matter.matcode) in ('" + singleMatter.ToString() + "')  ";

                DataSet myCliMat = _jurisUtility.RecordsetFromSQL(getCliMat);

                dataGridView1.DataSource = myCliMat.Tables[0];
            }

        }
    }
    
}
