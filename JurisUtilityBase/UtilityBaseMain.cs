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

        public string fromAtty = "";

        public string toAtty = "";

        private string typeOfTkpr = "";

        private List<string> matterIDs = new List<string>();

        private List<string> clientIDs = new List<string>();

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
                ///GetFieldLengths();
            }

            string TkprIndex;
            cbFrom.ClearItems();
            string SQLTkpr = "select empid + case when len(empid)=1 then '     ' when len(empid)=2 then '     ' when len(empid)=3 then '   ' else '  ' end +  empname as employee from employee where empvalidastkpr='Y' and (empsysnbr in (select distinct billtobillingatty from billto) or empsysnbr in (select distinct morigatty from matorigatty) or empsysnbr in (select distinct clibillingatty from client) or empsysnbr in (select distinct corigatty from cliorigatty)) order by empid";
            DataSet myRSTkpr = _jurisUtility.RecordsetFromSQL(SQLTkpr);

            if (myRSTkpr.Tables[0].Rows.Count == 0)
                cbFrom.SelectedIndex = 0;
            else
            {
                foreach (DataTable table in myRSTkpr.Tables)
                {

                    foreach (DataRow dr in table.Rows)
                    {
                        TkprIndex = dr["employee"].ToString();
                        cbFrom.Items.Add(TkprIndex);
                    }
                }

            }

            string TkprIndex2;
            cbTo.ClearItems();
            string SQLTkpr2 = "select empid + case when len(empid)=1 then '     ' when len(empid)=2 then '     ' when len(empid)=3 then '   ' else '  ' end +  empname as employee from employee where empvalidastkpr='Y' order by empid";
            DataSet myRSTkpr2 = _jurisUtility.RecordsetFromSQL(SQLTkpr2);


            if (myRSTkpr2.Tables[0].Rows.Count == 0)
                cbTo.SelectedIndex = 0;
            else
            {
                foreach (DataTable table in myRSTkpr2.Tables)
                {

                    foreach (DataRow dr in table.Rows)
                    {
                        TkprIndex2 = dr["employee"].ToString();
                        cbTo.Items.Add(TkprIndex2);
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
            int TkprSys = 0;
            string Tkpr = cbFrom.SelectedItem.ToString();
            int TkprLen = Tkpr.IndexOf(" ");
            string TkprSel = Tkpr.Substring(0, TkprLen);
            int TkprLength = TkprSel.Length;
            clientIDs.Clear();
            matterIDs.Clear();

            if (TkprLength < 1)
            { MessageBox.Show("Invalid Timekeeper. Please check your selections and try again."); }
            else
            {
                string SQLTkpr = "select empsysnbr as Tkpr from employee where empid = '" + TkprSel.ToString().Trim() + "'";
                DataSet myRSTkpr = _jurisUtility.RecordsetFromSQL(SQLTkpr);
                if (myRSTkpr.Tables[0].Rows.Count == 0)
                { MessageBox.Show("Invalid Timekeeper. Please check your selections and try again."); }
                else
                { TkprSys = Int32.Parse(myRSTkpr.Tables[0].Rows[0]["Tkpr"].ToString()); }


                int TkprSys2 = 0;
                string Tkpr2 = cbTo.SelectedItem.ToString();
                int TkprLen2 = Tkpr2.IndexOf(" ");
                string TkprSel2 = Tkpr2.Substring(0, TkprLen2);
                int TkprLength2 = TkprSel2.Length;
                DialogResult rsBoth;
                MessageBoxButtons buttons = MessageBoxButtons.YesNo;



                if (TkprLength2 < 1)
                { MessageBox.Show("Invalid Timekeeper. Please check your selections and try again."); }
                else
                {
                    string SQLTkpr2 = "select empsysnbr as Tkpr from employee where empid = '" + TkprSel2.ToString().Trim() + "'";
                    DataSet myRSTkpr2 = _jurisUtility.RecordsetFromSQL(SQLTkpr2);
                    if (myRSTkpr2.Tables[0].Rows.Count == 0)
                    { MessageBox.Show("Invalid Timekeeper. Please check your selections and try again."); }
                    else
                    { TkprSys2 = Int32.Parse(myRSTkpr2.Tables[0].Rows[0]["Tkpr"].ToString()); }

                }

                string BT = "";
                string CM2 = "";

                if (rbAll.Checked)
                { BT = "All Timekeepers for "; }
                if (rbTkprBill.Checked)
                { BT = "Billing Timekeepers for "; }
                if (rbTkprOrig.Checked)
                { BT = "Originating Timekeepers for "; }
                if (rbResp.Checked)
                { BT = "Responsible Timekeepers for "; }

                if (rbCM.Checked)
                { CM2 = " Clients and Matters will be changed from "; }
                if (rbClient.Checked)
                { CM2 = "Clients will be changed from "; }
                if (rbMatter.Checked)
                { CM2 = "Matters will be changed from "; }


                string RST = BT.ToString() + "" + CM2.ToString() + "" + TkprSel.ToString().Trim() + " to " + TkprSel2.ToString().Trim();

                rsBoth = MessageBox.Show(RST.ToString() + ". Do you wish to continue?", "Timekeeper Reassignment", buttons);

                if (rsBoth == DialogResult.No)
                { MessageBox.Show("Please check selections and try again."); }
                else
                {
                    if (rbAll.Checked)
                    {
                        typeOfTkpr = "Billing, Originating and Responsible ";
                        if (rbCM.Checked)
                        {
                            string CM = "";
                            string CB = "";
                            string SQLCM = "select convert(varchar(10),count(corigatty),1) as CO from cliorigatty where corigatty=" + TkprSys.ToString();
                            DataSet myRSCM = _jurisUtility.RecordsetFromSQL(SQLCM);
                            if (myRSCM.Tables[0].Rows.Count == 0)
                            { CM = "0"; }
                            else
                            { CM = myRSCM.Tables[0].Rows[0]["CO"].ToString(); }

                            string SQLCB = "select convert(varchar(10),count(clibillingatty),1) as CO from client where clibillingatty=" + TkprSys.ToString();
                            DataSet myRSCB = _jurisUtility.RecordsetFromSQL(SQLCB);
                            if (myRSCB.Tables[0].Rows.Count == 0)
                            { CB = "0"; }
                            else
                            { CB = myRSCB.Tables[0].Rows[0]["CO"].ToString(); }

                            OCTkpr(TkprSys, TkprSys2);

                            OMTkpr(TkprSys, TkprSys2);

                            BCTkpr(TkprSys, TkprSys2);

                            BMTkpr(TkprSys, TkprSys2);

                            RMTkpr(TkprSys, TkprSys2);

                            RCTkpr(TkprSys, TkprSys2);

                            Cursor.Current = Cursors.Default;
                            Application.DoEvents();
                            toolStripStatusLabel.Text = "Timekeepers Updated: Client/Matter Originating/Billing/Responsible " + CM.ToString() + "/" + CB.ToString();
                            statusStrip.Refresh();

                        }

                        if (rbClient.Checked)
                        {
                            string CM = "";
                            string CB = "";

                            string SQLCM = "select convert(varchar(10),count(corigatty),1) as CO from cliorigatty where corigatty=" + TkprSys.ToString();
                            DataSet myRSCM = _jurisUtility.RecordsetFromSQL(SQLCM);
                            if (myRSCM.Tables[0].Rows.Count == 0)
                            { CM = "0"; }
                            else
                            { CM = myRSCM.Tables[0].Rows[0]["CO"].ToString(); }

                            string SQLCB = "select convert(varchar(10),count(clibillingatty),1) as CO from client where clibillingatty=" + TkprSys.ToString();
                            DataSet myRSCB = _jurisUtility.RecordsetFromSQL(SQLCB);
                            if (myRSCB.Tables[0].Rows.Count == 0)
                            { CB = "0"; }
                            else
                            { CB = myRSCB.Tables[0].Rows[0]["CO"].ToString(); }


                            OCTkpr(TkprSys, TkprSys2);

                            BCTkpr(TkprSys, TkprSys2);

                            RCTkpr(TkprSys, TkprSys2);

                            Cursor.Current = Cursors.Default;
                            Application.DoEvents();
                            toolStripStatusLabel.Text = "Timekeepers Updated: Client Originating/Billing " + CM.ToString() + "/" + CB.ToString();
                            statusStrip.Refresh();

                        }

                        if (rbMatter.Checked)
                        {
                            string MO = "";
                            string MB = "";


                            string SQLMO = "select convert(varchar(10),count(morigatty),1) as CO from matorigatty where morigatty=" + TkprSys.ToString();
                            DataSet myRSMO = _jurisUtility.RecordsetFromSQL(SQLMO);
                            if (myRSMO.Tables[0].Rows.Count == 0)
                            { MO = "0"; }
                            else
                            { MO = myRSMO.Tables[0].Rows[0]["CO"].ToString(); }

                            string SQLMB = "select convert(varchar(10),count(billtobillingatty),1) as CO from billto where billtobillingatty=" + TkprSys.ToString();
                            DataSet myRSMB = _jurisUtility.RecordsetFromSQL(SQLMB);
                            if (myRSMB.Tables[0].Rows.Count == 0)
                            { MB = "0"; }
                            else
                            { MB = myRSMB.Tables[0].Rows[0]["CO"].ToString(); }

                            OMTkpr(TkprSys, TkprSys2);

                            BMTkpr(TkprSys, TkprSys2);

                            RMTkpr(TkprSys, TkprSys2);

                            Cursor.Current = Cursors.Default;
                            Application.DoEvents();
                            toolStripStatusLabel.Text = "Timekeepers Updated: Matter Originating/Billing " + MO.ToString() + "/" + MB.ToString();
                            statusStrip.Refresh();

                        }


                    }


                    if (rbTkprBill.Checked)
                    {
                        typeOfTkpr = "Billing ";
                        if (rbCM.Checked)
                        {
                            string CB = "";
                            string MB = "";

                            string SQLCB = "select convert(varchar(10),count(clibillingatty),1) as CO from client where clibillingatty=" + TkprSys.ToString();
                            DataSet myRSCB = _jurisUtility.RecordsetFromSQL(SQLCB);
                            if (myRSCB.Tables[0].Rows.Count == 0)
                            { CB = "0"; }
                            else
                            { CB = myRSCB.Tables[0].Rows[0]["CO"].ToString(); }


                            string SQLMB = "select convert(varchar(10),count(billtobillingatty),1) as CO from billto where billtobillingatty=" + TkprSys.ToString();
                            DataSet myRSMB = _jurisUtility.RecordsetFromSQL(SQLMB);
                            if (myRSMB.Tables[0].Rows.Count == 0)
                            { MB = "0"; }
                            else
                            { MB = myRSMB.Tables[0].Rows[0]["CO"].ToString(); }


                            BCTkpr(TkprSys, TkprSys2);
                            BMTkpr(TkprSys, TkprSys2);

                            Cursor.Current = Cursors.Default;
                            Application.DoEvents();
                            toolStripStatusLabel.Text = "Timekeepers Updated: Client Billing " + CB.ToString() + "; Matter Billing " + MB.ToString();
                            statusStrip.Refresh();

                        }

                        else if (rbClient.Checked)
                        {
                            string CB = "";
                            string SQLCB = "select convert(varchar(10),count(clibillingatty),1) as CO from client where clibillingatty=" + TkprSys.ToString();
                            DataSet myRSCB = _jurisUtility.RecordsetFromSQL(SQLCB);
                            if (myRSCB.Tables[0].Rows.Count == 0)
                            { CB = "0"; }
                            else
                            { CB = myRSCB.Tables[0].Rows[0]["CO"].ToString(); }

                            BCTkpr(TkprSys, TkprSys2);

                            Cursor.Current = Cursors.Default;
                            Application.DoEvents();
                            toolStripStatusLabel.Text = "Timekeepers Updated: Client Billing " + CB.ToString();
                            statusStrip.Refresh();
                        }

                        else if (rbMatter.Checked)
                        {
                            string MB = "";
                            string SQLMB = "select convert(varchar(10),count(billtobillingatty),1) as CO from billto where billtobillingatty=" + TkprSys.ToString();
                            DataSet myRSMB = _jurisUtility.RecordsetFromSQL(SQLMB);
                            if (myRSMB.Tables[0].Rows.Count == 0)
                            { MB = "0"; }
                            else
                            { MB = myRSMB.Tables[0].Rows[0]["CO"].ToString(); }

                            BMTkpr(TkprSys, TkprSys2);

                            Cursor.Current = Cursors.Default;
                            Application.DoEvents();
                            toolStripStatusLabel.Text = "Timekeepers Updated: Matter Billing " + MB.ToString();
                            statusStrip.Refresh();
                        }


                    }

                    if (rbTkprOrig.Checked)
                    {
                        typeOfTkpr = "Originating ";
                        if (rbCM.Checked)
                        {
                            string CM = "";
                            string MO = "";

                            string SQLCM = "select convert(varchar(10),count(corigatty),1) as CO from cliorigatty where corigatty=" + TkprSys.ToString();
                            DataSet myRSCM = _jurisUtility.RecordsetFromSQL(SQLCM);
                            if (myRSCM.Tables[0].Rows.Count == 0)
                            { CM = "0"; }
                            else
                            { CM = myRSCM.Tables[0].Rows[0]["CO"].ToString(); }



                            string SQLMO = "select convert(varchar(10),count(morigatty),1) as CO from matorigatty where morigatty=" + TkprSys.ToString();
                            DataSet myRSMO = _jurisUtility.RecordsetFromSQL(SQLMO);
                            if (myRSMO.Tables[0].Rows.Count == 0)
                            { MO = "0"; }
                            else
                            { MO = myRSMO.Tables[0].Rows[0]["CO"].ToString(); }



                            OCTkpr(TkprSys, TkprSys2);



                            OMTkpr(TkprSys, TkprSys2);



                            Cursor.Current = Cursors.Default;
                            Application.DoEvents();
                            toolStripStatusLabel.Text = "Timekeepers Updated: Client Originating " + CM.ToString() + "; Matter Originating " + MO.ToString();
                            statusStrip.Refresh();

                        }

                        if (rbClient.Checked)
                        {
                            string CM = "";
                            string SQLCM = "select convert(varchar(10),count(corigatty),1) as CO from cliorigatty where corigatty=" + TkprSys.ToString();
                            DataSet myRSCM = _jurisUtility.RecordsetFromSQL(SQLCM);
                            if (myRSCM.Tables[0].Rows.Count == 0)
                            { CM = "0"; }
                            else
                            { CM = myRSCM.Tables[0].Rows[0]["CO"].ToString(); }


                            OCTkpr(TkprSys, TkprSys2);


                            Cursor.Current = Cursors.Default;
                            Application.DoEvents();
                            toolStripStatusLabel.Text = "Timekeepers Updated: Client Originating " + CM.ToString();
                            statusStrip.Refresh();

                        }

                        if (rbMatter.Checked)
                        {
                            string MO = "";
                            string SQLMO = "select convert(varchar(10),count(morigatty),1) as CO from matorigatty where morigatty=" + TkprSys.ToString();
                            DataSet myRSMO = _jurisUtility.RecordsetFromSQL(SQLMO);
                            if (myRSMO.Tables[0].Rows.Count == 0)
                            { MO = "0"; }
                            else
                            { MO = myRSMO.Tables[0].Rows[0]["CO"].ToString(); }

                            OMTkpr(TkprSys, TkprSys2);


                            Cursor.Current = Cursors.Default;
                            Application.DoEvents();
                            toolStripStatusLabel.Text = "Timekeepers Updated: Matter Originating " + MO.ToString();
                            statusStrip.Refresh();
                        }

                    }
                    if (rbResp.Checked)
                    {
                        typeOfTkpr = "Responsible ";
                        if (rbCM.Checked)
                        {
                            string MO = "";
                            string MB = "";
                            string SQLMO = "select convert(varchar(10),count(mrtemployeeid),1) as CO from MatterResponsibleTimekeeper where mrtemployeeid=" + TkprSys.ToString();
                            DataSet myRSMO = _jurisUtility.RecordsetFromSQL(SQLMO);
                            if (myRSMO.Tables[0].Rows.Count == 0)
                            { MO = "0"; }
                            else
                            { MO = myRSMO.Tables[0].Rows[0]["CO"].ToString(); }

                            string SQLMB = "select convert(varchar(10),count(crtemployeeid),1) as CO from ClientResponsibleTimekeeper where crtemployeeid=" + TkprSys.ToString();
                            DataSet myRSMB = _jurisUtility.RecordsetFromSQL(SQLMB);
                            if (myRSMB.Tables[0].Rows.Count == 0)
                            { MB = "0"; }
                            else
                            { MB = myRSMB.Tables[0].Rows[0]["CO"].ToString(); }



                            RMTkpr(TkprSys, TkprSys2);

                            RCTkpr(TkprSys, TkprSys2);

                            Cursor.Current = Cursors.Default;
                            Application.DoEvents();
                            toolStripStatusLabel.Text = "Timekeepers Updated: Responsible Client/Matter " + MO.ToString() + "/" + MB.ToString();
                            statusStrip.Refresh();

                        }

                        if (rbClient.Checked)
                        {
                            string CM = "";
                            string SQLCM = "select convert(varchar(10),count(crtemployeeid),1) as CO from ClientResponsibleTimekeeper where crtemployeeid=" + TkprSys.ToString();
                            DataSet myRSCM = _jurisUtility.RecordsetFromSQL(SQLCM);
                            if (myRSCM.Tables[0].Rows.Count == 0)
                            { CM = "0"; }
                            else
                            { CM = myRSCM.Tables[0].Rows[0]["CO"].ToString(); }


                            RCTkpr(TkprSys, TkprSys2);


                            Cursor.Current = Cursors.Default;
                            Application.DoEvents();
                            toolStripStatusLabel.Text = "Timekeepers Updated: Client Responsible " + CM.ToString();
                            statusStrip.Refresh();

                        }

                        if (rbMatter.Checked)
                        {
                            string MO = "";
                            string SQLMO = "select convert(varchar(10),count(mrtemployeeid),1) as CO fromMatterResponsibleTimekeeper where mrtemployeeid=" + TkprSys.ToString();
                            DataSet myRSMO = _jurisUtility.RecordsetFromSQL(SQLMO);
                            if (myRSMO.Tables[0].Rows.Count == 0)
                            { MO = "0"; }
                            else
                            { MO = myRSMO.Tables[0].Rows[0]["CO"].ToString(); }

                            RMTkpr(TkprSys, TkprSys2);


                            Cursor.Current = Cursors.Default;
                            Application.DoEvents();
                            toolStripStatusLabel.Text = "Timekeepers Updated: Matter Responsible " + MO.ToString();
                            statusStrip.Refresh();
                        }

                    }

                    //add note
                    if (radioButtonNoteCardYes.Checked) //they selected YES
                    {
                        DateTime dt = new DateTime();
                        dt = DateTime.Today;
                        if (matterIDs.Any())
                        {
                            //update matter note card
                            List<string> currentList = matterIDs.Distinct().ToList();

                            foreach (string id in currentList)
                            {
                                string noteIndex = "TkprChg";
                                int results = findHowManyNotesAlreadySaytkprChange(true, id);
                                if (results > 0)
                                    noteIndex = noteIndex + results.ToString();
                                string CC2 = "insert into [MatterNote] ([MNMatter],[MNNoteIndex],[MNObject],[MNNoteText],[MNNoteObject]) values (" + id + ", '" + noteIndex + "', '', 'Updated Timekeeper: " + typeOfTkpr + " changed from " + fromAtty + " to " + toAtty + " on " + dt.ToShortDateString() + "', null)";
                                _jurisUtility.ExecuteNonQueryCommand(0, CC2);
                            }
                        }
                        if (clientIDs.Any())
                        {
                            //update client note card
                            List<string> currentList = clientIDs.Distinct().ToList();

                            foreach (string id in currentList)
                            {
                                string noteIndex = "TkprChg";
                                int results = findHowManyNotesAlreadySaytkprChange(false, id);
                                if (results != 0)
                                    noteIndex = noteIndex + results.ToString();
                                string CC2 = "insert into [ClientNote] ([CNClient],[CNNoteIndex],[CNObject],[CNNoteText],[CNNoteObject]) values (" + id + ", '" + noteIndex + "', '', 'Updated Timekeeper: " + typeOfTkpr + " changed from " + fromAtty + " to " + toAtty + " on " + dt.ToShortDateString() + "', null)";
                                _jurisUtility.ExecuteNonQueryCommand(0, CC2);

                            }
                        }
                    }
                    typeOfTkpr = "";
                    UpdateStatus("Timekeeper Update Complete", 5, 5);
                    string LogNote = "Timekeeper Assignment - " + TkprSel.ToString().Trim() + " to " + TkprSel2.ToString().Trim();
                    WriteLog(LogNote.ToString());
                    //outputs the post report
                    var result = MessageBox.Show("Process Complete! Would you like to see the post log?", "Log confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (result == DialogResult.Yes)
                    {
                        string reportSQL = getPostReportSQL();

                        DataSet ds = _jurisUtility.RecordsetFromSQL(reportSQL);


                        ReportDisplay rpds = new ReportDisplay(ds);
                        rpds.Show();
                    }

                }


            }


        }
    

        private int findHowManyNotesAlreadySaytkprChange(bool isNoteMatter, string ID) //true is for matterNote, false is for clientNotes
        {
            DataSet ds = new DataSet();
            string sql = "";
            if (isNoteMatter)
                sql = "select MNMatter from MatterNote where MNNoteIndex like 'TkprChg%' and MNMatter = " + ID;
            else
                sql = "select CNClient from ClientNote where CNNoteIndex like 'TkprChg%' and CNClient = " + ID;
            ds = _jurisUtility.RecordsetFromSQL(sql);
            if (ds == null || ds.Tables.Count == 0 || ds.Tables[0].Rows.Count == 0)
                return 0;
            else
                return ds.Tables[0].Rows.Count;
        }



        
        private void OCTkpr(int TkFrom, int TkTo)
        {  
            Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();
            toolStripStatusLabel.Text = "Updating Client Originating Timekeepers....";
            statusStrip.Refresh();
            UpdateStatus("Client Originating Timekeepers", 1, 5);


            string SQLCO = @"select cast(corigcli as varchar(10)) as Cli, cast(corigatty as varchar(10)) as Corig, cast(corigpcnt as varchar(10)) as Pct from cliorigatty where corigatty=cast('" + TkFrom.ToString() + "' as int)";
            DataSet myRSCMQ = _jurisUtility.RecordsetFromSQL(SQLCO);
            if (myRSCMQ.Tables[0].Rows.Count != 0)

            {

                foreach (DataRow row in myRSCMQ.Tables[0].Rows)
                {
                    string Cli = row["Cli"].ToString();
                    clientIDs.Add(Cli);
                    string OAtty = row["Corig"].ToString();
                    string Pct = row["Pct"].ToString();

                    string SQLCT = @"select cast(corigcli as varchar(10)) as Cli, cast(corigatty as varchar(10)) as Corig, cast(corigpcnt as varchar(10)) as Pct from cliorigatty where corigatty=cast('" + TkTo.ToString() + "' as int) and corigcli=cast('" + Cli.ToString() + "' as int)";
                    DataSet myRSCZ = _jurisUtility.RecordsetFromSQL(SQLCT);

                    if (myRSCZ.Tables[0].Rows.Count != 0)
                    { 
                        string CC = "delete from cliorigatty  where corigcli=cast('" + Cli.ToString() + "' as int) and corigatty=cast('" + TkFrom.ToString() + "' as int)";
                        _jurisUtility.ExecuteNonQueryCommand(0, CC);

                        string CC2 = "update cliorigatty set corigpcnt = corigpcnt + cast('" + Pct.ToString() + "' as decimal(7,4))  where corigcli=cast('" + Cli.ToString() + "' as int) and corigatty=cast('" + TkTo.ToString() + "' as int)";
                        _jurisUtility.ExecuteNonQueryCommand(0, CC2);
                        myRSCZ.Clear();
                    }
                    else
                    { string CC3 = "update cliorigatty set corigatty=" + TkTo.ToString() + " where corigcli=cast('" + Cli.ToString() + "' as int) and corigatty=cast('" + TkFrom.ToString() + "' as int)";
                        _jurisUtility.ExecuteNonQueryCommand(0, CC3);
                        myRSCZ.Clear();
                    }

                }
               }

           }

        private void OMTkpr(int TkFrom, int TkTo)
        {
            Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();
            toolStripStatusLabel.Text = "Updating Matter Originating Timekeepers....";
            statusStrip.Refresh();
            UpdateStatus("Matter Originating Timekeepers", 2, 5);
            string SQLCO = @"select cast(morigmat as varchar(10)) as Cli, cast(morigatty as varchar(10)) as Corig, cast(morigpcnt as varchar(10)) as Pct from matorigatty where morigatty=" + TkFrom.ToString();
            DataSet myRSCY = _jurisUtility.RecordsetFromSQL(SQLCO);
            if (myRSCY.Tables[0].Rows.Count != 0)

            {

                foreach (DataRow row in myRSCY.Tables[0].Rows)

                {
                    string Cli = row["Cli"].ToString();
                    matterIDs.Add(Cli);
                    string OAtty = row["Corig"].ToString();
                    string Pct = row["Pct"].ToString();

                    string SQLCZM = @"select cast(morigmat as varchar(10)) as Cli, cast(morigatty as varchar(10)) as Corig, cast(morigpcnt as varchar(10)) as Pct from matorigatty where morigatty=cast('" + TkTo.ToString() + "' as int) and morigmat=cast('" + Cli.ToString() + "' as int)";
                    DataSet myRSCX = _jurisUtility.RecordsetFromSQL(SQLCZM);

                    if (myRSCX.Tables[0].Rows.Count != 0)
                    {

                        string CC = "delete from matorigatty  where morigmat=" + Cli.ToString() + " and morigatty=" + TkFrom.ToString();
                        _jurisUtility.ExecuteNonQueryCommand(0, CC);

                        string CC2 = "update matorigatty set morigpcnt = morigpcnt + cast('" + Pct.ToString() + "' as decimal(7,4))  where morigmat=cast('" + Cli.ToString() + "' as int) and  morigatty=cast('" + TkTo.ToString() + "' as int)";
                        _jurisUtility.ExecuteNonQueryCommand(0, CC2);

                        myRSCX.Clear();


                    }
                    else
                    {
                        string CC = "update matorigatty set morigatty=" + TkTo.ToString() + " where morigmat=cast('" + Cli.ToString() + "' as int) and morigatty=cast('" + TkFrom.ToString() + "' as int)";
                        _jurisUtility.ExecuteNonQueryCommand(0, CC);
                        myRSCX.Clear();

                    }


                }
            }
        }


        private void BMTkpr(int TkFrom, int TkTo)
        {

            Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();
            toolStripStatusLabel.Text = "Updating Matter Billing Timekeepers....";
            statusStrip.Refresh();
            UpdateStatus("Matter Billing Timekeepers", 4, 5);
            string SQLCO = @"select distinct cast(MatSysNbr as varchar(10)) as Cli from BillTo inner join matter on matter.matbillto = billto.BillToSysNbr where billtobillingatty=" + TkFrom.ToString();
            DataSet myRSCY = _jurisUtility.RecordsetFromSQL(SQLCO);
            if (myRSCY.Tables[0].Rows.Count != 0)
            {

                foreach (DataRow row in myRSCY.Tables[0].Rows)
                {
                    string Cli = row["Cli"].ToString();
                    matterIDs.Add(Cli);
                }
            }

            string SQLMB = @"update billto set billtobillingatty=" + TkTo.ToString() + " where billtobillingatty=" + TkFrom.ToString();
            _jurisUtility.ExecuteNonQueryCommand(0, SQLMB);

        }


        private void BCTkpr(int TkFrom, int TkTo)
        {
            Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();
            toolStripStatusLabel.Text = "Updating Client Billing Timekeepers....";
            statusStrip.Refresh();
            UpdateStatus("Client Billing Timekeepers", 3, 5);
            string SQLCO = @"select distinct cast(cliSysNbr as varchar(10)) as Cli from client where clibillingatty=" + TkFrom.ToString();
            DataSet myRSCY = _jurisUtility.RecordsetFromSQL(SQLCO);
            if (myRSCY.Tables[0].Rows.Count != 0)
            {

                foreach (DataRow row in myRSCY.Tables[0].Rows)
                {
                    string Cli = row["Cli"].ToString();
                    clientIDs.Add(Cli);
                }
            }


            string SQLCB = @"update client set clibillingatty=" + TkTo.ToString() + " where clibillingatty=" + TkFrom.ToString();
            _jurisUtility.ExecuteNonQueryCommand(0, SQLCB);

        }


        private void RCTkpr(int TkFrom, int TkTo)
        {
            Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();
            toolStripStatusLabel.Text = "Updating Client Responsible Timekeepers....";
            statusStrip.Refresh();
            UpdateStatus("Client Responsible Timekeepers", 1, 5);


            string SQLCO = @"select cast(crtclientid as varchar(10)) as Cli, cast(crtemployeeid as varchar(10)) as Corig, cast(crtpercent as varchar(10)) as Pct from ClientResponsibleTimekeeper where crtemployeeid=cast('" + TkFrom.ToString() + "' as int)";
            DataSet myRSCMQ = _jurisUtility.RecordsetFromSQL(SQLCO);
            if (myRSCMQ.Tables[0].Rows.Count != 0)
            {

                foreach (DataRow row in myRSCMQ.Tables[0].Rows)
                {

                    string Cli = row["Cli"].ToString();
                    clientIDs.Add(Cli);
                    string OAtty = row["Corig"].ToString();
                    string Pct = row["Pct"].ToString();

                    string SQLCT = @"select cast(crtclientid as varchar(10)) as Cli, cast(crtemployeeid as varchar(10)) as Corig, cast(crtpercent as varchar(10)) as Pct from ClientResponsibleTimekeeper where crtemployeeid=cast('" + TkTo.ToString() + "' as int) and crtclientid=cast('" + Cli.ToString() + "' as int)";
                    DataSet myRSCZ = _jurisUtility.RecordsetFromSQL(SQLCT);

                    if (myRSCZ.Tables[0].Rows.Count != 0)
                    {
                        string CC = "delete from ClientResponsibleTimekeeper  where crtclientid=cast('" + Cli.ToString() + "' as int) and crtemployeeid=cast('" + TkFrom.ToString() + "' as int)";
                        _jurisUtility.ExecuteNonQueryCommand(0, CC);

                        string CC2 = "update ClientResponsibleTimekeeper set CRTPercent = crtpercent + cast('" + Pct.ToString() + "' as decimal(7,4))  where crtclientid=cast('" + Cli.ToString() + "' as int) and crtemployeeid=cast('" + TkTo.ToString() + "' as int)";
                        _jurisUtility.ExecuteNonQueryCommand(0, CC2);
                        myRSCZ.Clear();
                    }
                    else
                    {
                        string CC3 = "update ClientResponsibleTimekeeper set crtemployeeid=" + TkTo.ToString() + " where crtclientid=cast('" + Cli.ToString() + "' as int) and crtemployeeid=cast('" + TkFrom.ToString() + "' as int)";
                        _jurisUtility.ExecuteNonQueryCommand(0, CC3);
                        myRSCZ.Clear();
                    }

                }
            }
        }

        private void RMTkpr(int TkFrom, int TkTo)
        {
            Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();
            toolStripStatusLabel.Text = "Updating Matter Responsible Timekeepers....";
            statusStrip.Refresh();
            UpdateStatus("Matter Responsible Timekeepers", 2, 5);
            string SQLCO = @"select cast(mrtmatterid as varchar(10)) as Cli, cast(MRTEmployeeID as varchar(10)) as Corig, cast(MRTPercent as varchar(10)) as Pct from MatterResponsibleTimekeeper where MRTEmployeeID=cast('" + TkFrom.ToString() + "' as int)";
            DataSet myRSCY = _jurisUtility.RecordsetFromSQL(SQLCO);
            if (myRSCY.Tables[0].Rows.Count != 0)
            {

                foreach (DataRow row in myRSCY.Tables[0].Rows)
                {
                    string Cli = row["Cli"].ToString();
                    matterIDs.Add(Cli);
                    string OAtty = row["Corig"].ToString();
                    string Pct = row["Pct"].ToString();


                    string SQLCZM = @"select cast(mrtmatterid as varchar(10)) as Cli, cast(mrtemployeeid as varchar(10)) as Corig, cast(mrtpercent as varchar(10)) as Pct from MatterResponsibleTimekeeper where mrtemployeeid=cast('" + TkTo.ToString() + "' as int) and mrtmatterid=cast('" + Cli.ToString() + "' as int)";
                    DataSet myRSCX = _jurisUtility.RecordsetFromSQL(SQLCZM);

                    if (myRSCX.Tables[0].Rows.Count != 0)
                    {

                        string CC = "delete from MatterResponsibleTimekeeper  where mrtmatterid=" + Cli.ToString() + " and mrtemployeeid=" + TkFrom.ToString();
                        _jurisUtility.ExecuteNonQueryCommand(0, CC);

                        string CC2 = "update MatterResponsibleTimekeeper set mrtpercent = mrtpercent + cast('" + Pct.ToString() + "' as decimal(7,4))  where mrtmatterid=cast('" + Cli.ToString() + "' as int) and  mrtemployeeid=cast('" + TkTo.ToString() + "' as int)";
                        _jurisUtility.ExecuteNonQueryCommand(0, CC2);

                        myRSCX.Clear();


                    }
                    else
                    {
                        string CC = "update MatterResponsibleTimekeeper set mrtemployeeid=" + TkTo.ToString() + " where mrtmatterid=cast('" + Cli.ToString() + "' as int) and mrtemployeeid=cast('" + TkFrom.ToString() + "' as int)";
                        _jurisUtility.ExecuteNonQueryCommand(0, CC);

                        myRSCX.Clear();
                    }


                }
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
            string sql = "Insert Into UtilityLog(ULTimeStamp,ULWkStaUser,ULComment) Values(convert(datetime,'" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "'),cast('" +  GetComputerAndUser() + "' as varchar(100))," + "cast('" + comment.ToString() + "' as varchar(8000)))";
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

        private void btExit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void cbFrom_SelectedIndexChanged(object sender, EventArgs e)
        {
            fromAtty = cbFrom.Text;
            fromAtty = fromAtty.Split(' ')[0];
        }

        private void cbTo_SelectedIndexChanged(object sender, EventArgs e)
        {
            toAtty = cbTo.Text;
            toAtty = toAtty.Split(' ')[0];
        }

        private void buttonReport_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(toAtty) || string.IsNullOrEmpty(fromAtty))
                MessageBox.Show("Please select from both Timekeeper drop downs", "Selection Error");
            else
            {
                //generates output of the report for before and after the change will be made to client
                string SQLTkpr = getReportSQL();

                DataSet ds = _jurisUtility.RecordsetFromSQL(SQLTkpr);

                ReportDisplay rpds = new ReportDisplay(ds);
                rpds.Show();

            }
        }

        private string getReportSQL()
        {
            string reportSQL = "";
            //if matter and billing timekeeper
            if (rbTkprBill.Checked && rbMatter.Checked)
                reportSQL = "select Clicode as ClientCode, Clireportingname as ClientName, Matcode as MatterCode, Matreportingname as MatterName,empid as CurrentBillingTimekeeper, '" + toAtty + "' as NewBillingTimekeeper" +
                        " from matter" +
                        " inner join client on matclinbr=clisysnbr" +
                        " inner join billto on matbillto=billtosysnbr" +
                        " inner join employee on empsysnbr=billtobillingatty" +
                        " where empid='" + fromAtty + "'";


            //if matter and originating timekeeper
            else if (rbMatter.Checked && rbTkprOrig.Checked)
                reportSQL = "select Clicode as ClientCode, Clireportingname as ClientName, Matcode as MatterCode, Matreportingname as MatterName,empid as CurrentOriginatingTimekeeper, '" + toAtty + "' as NewOriginatingTimekeeper" +
                    " from matter" +
                    " inner join client on matclinbr=clisysnbr" +
                    " inner join matorigatty on matsysnbr=morigmat" +
                    " inner join employee on empsysnbr=morigatty" +
                    " where empid='" + fromAtty + "'";

            //if matter and responsible
            else if (rbMatter.Checked && rbResp.Checked)
                reportSQL = "select Clicode as ClientCode, Clireportingname as ClientName, Matcode as MatterCode, Matreportingname as MatterName,mrtemployeeid as CurrentResponsibleTimekeeper, '" + toAtty + "' as NewResponsibleTimekeeper" +
                    " from matter" +
                    " inner join client on matclinbr=clisysnbr" +
                    " inner join MatterResponsibleTimekeeper on matsysnbr=mrtmatterid" +
                    " inner join employee on empsysnbr=mrtemployeeid" +
                    " where empid='" + fromAtty + "'";

            //if client and responsible
            else if (rbClient.Checked && rbResp.Checked)
                reportSQL = "select Clicode as ClientCode, Clireportingname as ClientName, Matcode as MatterCode, Matreportingname as MatterName,crtemployeeid as CurrentResponsibleTimekeeper, '" + toAtty + "' as NewResponsibleTimekeeper" +
                    " from matter" +
                    " inner join client on matclinbr=clisysnbr" +
                    " inner join ClientResponsibleTimekeeper on CliSysNbr=crtclientid" +
                    " inner join employee on empsysnbr=crtemployeeid" +
                    " where empid='" + fromAtty + "'";

            //if client and billing timekeeper
            else if (rbTkprBill.Checked && rbClient.Checked)
                reportSQL = "select Clicode as ClientCode, Clireportingname as ClientName, empid as CurrentBillingTimekeeper,   '" + toAtty + "' as NewBillingTimekeeper" +
                    " from  client" +
                    " inner join employee on empsysnbr=clibillingatty" +
                    " where empid='" + fromAtty + "'";

            //if client  and originating timekeeper
            else if (rbClient.Checked && rbTkprOrig.Checked)
                reportSQL = "select Clicode as ClientCode, Clireportingname as ClientName, empid as CurrentOriginatingTimekeeper,  '" + toAtty + "' as NewOriginatingTimekeeper" +
                    " from  client" +
                    " inner join cliorigatty on corigcli=clisysnbr" +
                    " inner join employee on empsysnbr=corigatty" +
                    " where empid='" + fromAtty + "'";


            //if client & matter....and billing
            else if (rbCM.Checked && rbTkprBill.Checked)
                reportSQL = "select Clicode as ClientCode, Clireportingname as ClientName, Matcode as MatterCode, Matreportingname as MatterName,matBill.empid as MatterBillingTimekeeper, cliBill.empid as ClientBillingTimekeeper, '" + toAtty + "' as NewBillingTimekeeper" +
                        " from matter" +
                        " inner join client on matclinbr=clisysnbr" +
                        " inner join billto on matbillto=billtosysnbr" +
                        " inner join employee matBill on matBill.empsysnbr=billto.billtobillingatty" +
                        " inner join employee cliBill on cliBill.empsysnbr=client.clibillingatty" +
                        " where cliBill.empid='" + fromAtty + "' or matBill.empid ='" + fromAtty + "'";

                //if client & matter....and originating
            else if (rbCM.Checked && rbTkprOrig.Checked)
                reportSQL = "select Clicode as ClientCode, Clireportingname as ClientName, Matcode as MatterCode, Matreportingname as MatterName,matOrig.empid as MatterOriginatingTimekeeper,  cliOrig.empid as ClientOriginatingTimekeeper, '" + toAtty + "' as NewOriginatingTimekeeper" +
                    " from matter" +
                    " inner join client on matclinbr=clisysnbr" +
                    " inner join matorigatty on matsysnbr=morigmat" +
                    " inner join cliorigatty on corigcli=clisysnbr" +
                    " inner join employee matOrig on matOrig.empsysnbr=morigatty" +
                    " inner join employee cliOrig on cliOrig.empsysnbr=corigatty" +
                    " where matOrig.empid='" + fromAtty + "' or cliOrig.empid ='" + fromAtty + "'";
                //if client & matter...and Responsible
            else if (rbCM.Checked && rbResp.Checked)
                reportSQL = "select Clicode as ClientCode, Clireportingname as ClientName, Matcode as MatterCode, Matreportingname as MatterName,matOrig.empid as MatterResponsibleTimekeeper,  cliOrig.empid as ClientResponsibleTimekeeper, '" + toAtty + "' as NewResponsibleTimekeeper" +
                    " from matter" +
                    " inner join client on matclinbr=clisysnbr" +
                    " inner join MatterResponsibleTimekeeper on matsysnbr=mrtmatterid" +
                    " inner join ClientResponsibleTimekeeper on crtclientid=clisysnbr" +
                    " inner join employee matOrig on matOrig.empsysnbr=mrtemployeeid" +
                    " inner join employee cliOrig on cliOrig.empsysnbr=crtemployeeid" +
                    " where matOrig.empid='" + fromAtty + "' or cliOrig.empid ='" + fromAtty + "'";

            //matter and all
            else if (rbMatter.Checked && rbAll.Checked)
                reportSQL = "select Clicode as ClientCode, Clireportingname as ClientName, Matcode as MatterCode, Matreportingname as MatterName,matOrig.empid as MatterOriginatingTimekeeper,  matBill.empid as MatterBillingTimekeeper,  matResp.empid as MatterResponsibleTimekeeper,  '" + toAtty + "' as NewTimekeeperForAll" +
                    " from matter" +
                    " inner join client on clisysnbr = matclinbr " + 
                    " inner join matorigatty on matsysnbr=morigmat" +
                    " inner join billto on matbillto=billtosysnbr" +
                    " left outer join MatterResponsibleTimekeeper on matsysnbr=mrtmatterid" +
                    " inner join employee matOrig on matOrig.empsysnbr=morigatty" +
                    " inner join employee matBill on matBill.empsysnbr=billto.billtobillingatty" +
                    " left outer join employee matResp on matResp.empsysnbr=mrtemployeeid" +
                    " where matOrig.empid='" + fromAtty + "'" +
                    " or matBill.empid='" + fromAtty + "'" +
                    " or matResp.empid='" + fromAtty + "'";

            //client and all
            else if (rbClient.Checked && rbAll.Checked)
                reportSQL = "select Clicode as ClientCode, Clireportingname as ClientName,  cliOrig.empid as ClientOriginatingTimekeeper,cliBill.empid as ClientBillingTimekeeper, cliResp.empid as ClientResponsibleTimekeeper, '" + toAtty + "' as NewTimekeeperForAll" +
                    " from client" +
                    " inner join cliorigatty on corigcli=clisysnbr" +
                    " left outer join ClientResponsibleTimekeeper on crtclientid=clisysnbr" +
                    " inner join employee cliOrig on cliOrig.empsysnbr=corigatty" +
                    " inner join employee cliBill on cliBill.empsysnbr=client.clibillingatty" +
                    " left outer join employee cliResp on cliResp.empsysnbr=crtemployeeid" +
                    " where cliOrig.empid ='" + fromAtty + "'" +
                    " or cliBill.empid ='" + fromAtty + "'" +
                    " or cliResp.empid ='" + fromAtty + "'";

            //if client & matter.....and All
            else if (rbCM.Checked && rbAll.Checked)
                reportSQL = "select Clicode as ClientCode, Clireportingname as ClientName, Matcode as MatterCode, Matreportingname as MatterName,matOrig.empid as MatterOriginatingTimekeeper,  cliOrig.empid as ClientOriginatingTimekeeper,matBill.empid as MatterBillingTimekeeper,  cliBill.empid as ClientBillingTimekeeper, matResp.empid as MatterResponsibleTimekeeper,  cliResp.empid as ClientResponsibleTimekeeper, '" + toAtty + "' as NewTimekeeperForAll" +
                    " from matter" +
                    " inner join client on matclinbr=clisysnbr" +
                    " inner join matorigatty on matsysnbr=morigmat" +
                    " inner join cliorigatty on corigcli=clisysnbr" +
                    " inner join billto on matbillto=billtosysnbr" +
                    " left outer join MatterResponsibleTimekeeper on matsysnbr=mrtmatterid" +
                    " left outer join ClientResponsibleTimekeeper on crtclientid=clisysnbr" +
                    " inner join employee matOrig on matOrig.empsysnbr=morigatty" +
                    " inner join employee cliOrig on cliOrig.empsysnbr=corigatty" +
                    " inner join employee matBill on matBill.empsysnbr=billto.billtobillingatty" +
                    " inner join employee cliBill on cliBill.empsysnbr=client.clibillingatty" +
                    " left outer join employee matResp on matResp.empsysnbr=mrtemployeeid" +
                    " left outer join employee cliResp on cliResp.empsysnbr=crtemployeeid" +
                    " where matOrig.empid='" + fromAtty + "' or cliOrig.empid ='" + fromAtty + "'" +
                    " or matBill.empid='" + fromAtty + "' or cliBill.empid ='" + fromAtty + "'" + 
                    " or matResp.empid='" + fromAtty + "' or cliResp.empid ='" + fromAtty + "'";

            return reportSQL;
        }

        private string getPostReportSQL()
        {
            string reportSQL = "";
            string clients = String.Join(",", clientIDs.Select(x => x.ToString()).ToArray());
            string matters = String.Join(",", matterIDs.Select(x => x.ToString()).ToArray());
            //if matter and billing timekeeper
            if (rbTkprBill.Checked && rbMatter.Checked)
                reportSQL = "select Clicode as ClientCode, Clireportingname as ClientName, Matcode as MatterCode, Matreportingname as MatterName,empid as CurrentBillingTimekeeper" +
                        " from matter" +
                        " inner join client on matclinbr=clisysnbr" +
                        " inner join billto on matbillto=billtosysnbr" +
                        " inner join employee on empsysnbr=billtobillingatty" +
                        " where matsysnbr in (" + matters + ")";


            //if matter and originating timekeeper
            else if (rbMatter.Checked && rbTkprOrig.Checked)
                reportSQL = "select Clicode as ClientCode, Clireportingname as ClientName, Matcode as MatterCode, Matreportingname as MatterName,empid as CurrentOriginatingTimekeeper" +
                    " from matter" +
                    " inner join client on matclinbr=clisysnbr" +
                    " inner join matorigatty on matsysnbr=morigmat" +
                    " inner join employee on empsysnbr=morigatty" +
                        " where matsysnbr in (" + matters + ")";

            //if matter and responsible
            else if (rbMatter.Checked && rbResp.Checked)
                reportSQL = "select Clicode as ClientCode, Clireportingname as ClientName, Matcode as MatterCode, Matreportingname as MatterName,mrtemployeeid as CurrentResponsibleTimekeeper" +
                    " from matter" +
                    " inner join client on matclinbr=clisysnbr" +
                    " inner join MatterResponsibleTimekeeper on matsysnbr=mrtmatterid" +
                    " inner join employee on empsysnbr=mrtemployeeid" +
                        " where matsysnbr in (" + matters + ")";

            //if client and responsible
            else if (rbClient.Checked && rbResp.Checked)
                reportSQL = "select Clicode as ClientCode, Clireportingname as ClientName, Matcode as MatterCode, Matreportingname as MatterName,crtemployeeid as CurrentResponsibleTimekeeper" +
                    " from matter" +
                    " inner join client on matclinbr=clisysnbr" +
                    " inner join ClientResponsibleTimekeeper on CliSysNbr=crtclientid" +
                    " inner join employee on empsysnbr=crtemployeeid" +
                    " where clisysnbr in (" + clients + ")";

            //if client and billing timekeeper
            else if (rbTkprBill.Checked && rbClient.Checked)
                reportSQL = "select Clicode as ClientCode, Clireportingname as ClientName, empid as CurrentBillingTimekeeper" +
                    " from  client" +
                    " inner join employee on empsysnbr=clibillingatty" +
                    " where clisysnbr in (" + clients + ")";

            //if client  and originating timekeeper
            else if (rbClient.Checked && rbTkprOrig.Checked)
                reportSQL = "select Clicode as ClientCode, Clireportingname as ClientName, empid as CurrentOriginatingTimekeeper" +
                    " from  client" +
                    " inner join cliorigatty on corigcli=clisysnbr" +
                    " inner join employee on empsysnbr=corigatty" +
                    " where clisysnbr in (" + clients + ")";


            //if client & matter....and billing
            else if (rbCM.Checked && rbTkprBill.Checked)
                reportSQL = "select Clicode as ClientCode, Clireportingname as ClientName, Matcode as MatterCode, Matreportingname as MatterName,matBill.empid as MatterBillingTimekeeper, cliBill.empid as ClientBillingTimekeeper" +
                        " from matter" +
                        " inner join client on matclinbr=clisysnbr" +
                        " inner join billto on matbillto=billtosysnbr" +
                        " inner join employee matBill on matBill.empsysnbr=billto.billtobillingatty" +
                        " inner join employee cliBill on cliBill.empsysnbr=client.clibillingatty" +
                        " where matsysnbr in (" + matters + ") or clisysnbr in (" + clients + ")";

            //if client & matter....and responsible
            else if (rbResp.Checked && rbTkprBill.Checked)
                reportSQL = "select Clicode as ClientCode, Clireportingname as ClientName, Matcode as MatterCode, Matreportingname as MatterName,matOrig.empid as MatterResponsibleTimekeeper,  cliOrig.empid as ClientResponsibleTimekeeper, '" + toAtty + "' as NewResponsibleTimekeeper" +
                    " from matter" +
                    " inner join client on matclinbr=clisysnbr" +
                    " inner join MatterResponsibleTimekeeper on matsysnbr=mrtmatterid" +
                    " inner join ClientResponsibleTimekeeper on crtclientid=clisysnbr" +
                    " inner join employee matOrig on matOrig.empsysnbr=mrtemployeeid" +
                        " where matsysnbr in (" + matters + ") or clisysnbr in (" + clients + ")";

            //if client & matter....and originating
            else if (rbCM.Checked && rbTkprOrig.Checked)
                reportSQL = "select Clicode as ClientCode, Clireportingname as ClientName, Matcode as MatterCode, Matreportingname as MatterName,matOrig.empid as MatterOriginatingTimekeeper,  cliOrig.empid as ClientOriginatingTimekeeper" +
                    " from matter" +
                    " inner join client on matclinbr=clisysnbr" +
                    " inner join matorigatty on matsysnbr=morigmat" +
                    " inner join cliorigatty on corigcli=clisysnbr" +
                    " inner join employee matOrig on matOrig.empsysnbr=morigatty" +
                    " inner join employee cliOrig on cliOrig.empsysnbr=corigatty" +
                        " where matsysnbr in (" + matters + ") or clisysnbr in (" + clients + ")";

            //if client & matter.....and All
            else if (rbCM.Checked && rbAll.Checked)
                reportSQL = "select Clicode as ClientCode, Clireportingname as ClientName, Matcode as MatterCode, Matreportingname as MatterName,matOrig.empid as MatterOriginatingTimekeeper,  cliOrig.empid as ClientOriginatingTimekeeper,matBill.empid as MatterBillingTimekeeper,  cliBill.empid as ClientBillingTimekeeper, matResp.empid as MatterResponsibleTimekeeper,  cliResp.empid as ClientResponsibleTimekeeper" +
                    " from matter" +
                    " inner join client on matclinbr=clisysnbr" +
                    " inner join matorigatty on matsysnbr=morigmat" +
                    " inner join cliorigatty on corigcli=clisysnbr" +
                    " inner join billto on matbillto=billtosysnbr" +
                    " left outer join MatterResponsibleTimekeeper on matsysnbr=mrtmatterid" +
                    " left outer join ClientResponsibleTimekeeper on crtclientid=clisysnbr" +
                    " inner join employee matOrig on matOrig.empsysnbr=morigatty" +
                    " inner join employee cliOrig on cliOrig.empsysnbr=corigatty" +
                    " inner join employee matBill on matBill.empsysnbr=billto.billtobillingatty" +
                    " inner join employee cliBill on cliBill.empsysnbr=client.clibillingatty" +
                    " left outer join employee matResp on matResp.empsysnbr=mrtemployeeid" +
                    " left outer join employee cliResp on cliResp.empsysnbr=crtemployeeid" +
                        " where matsysnbr in (" + matters + ") or clisysnbr in (" + clients + ")";

            //matter and all
            else if (rbMatter.Checked && rbAll.Checked)
                reportSQL = "select Clicode as ClientCode, Clireportingname as ClientName, Matcode as MatterCode, Matreportingname as MatterName,matOrig.empid as MatterOriginatingTimekeeper,  matBill.empid as MatterBillingTimekeeper,  matResp.empid as MatterResponsibleTimekeeper,  '" + toAtty + "' as NewTimekeeperForAll" +
                    " from matter" +
                    " inner join client on clisysnbr = matclinbr " +
                    " inner join matorigatty on matsysnbr=morigmat" +
                    " inner join billto on matbillto=billtosysnbr" +
                    " left outer join MatterResponsibleTimekeeper on matsysnbr=mrtmatterid" +
                    " inner join employee matOrig on matOrig.empsysnbr=morigatty" +
                    " inner join employee matBill on matBill.empsysnbr=billto.billtobillingatty" +
                    " left outer join employee matResp on matResp.empsysnbr=mrtemployeeid" +
                        " where matsysnbr in (" + matters + ")";

            //client and all
            else if (rbClient.Checked && rbAll.Checked)
                reportSQL = "select Clicode as ClientCode, Clireportingname as ClientName,  cliOrig.empid as ClientOriginatingTimekeeper,cliBill.empid as ClientBillingTimekeeper, cliResp.empid as ClientResponsibleTimekeeper, '" + toAtty + "' as NewTimekeeperForAll" +
                    " from client" +
                    " inner join cliorigatty on corigcli=clisysnbr" +
                    " left outer join ClientResponsibleTimekeeper on crtclientid=clisysnbr" +
                    " inner join employee cliOrig on cliOrig.empsysnbr=corigatty" +
                    " inner join employee cliBill on cliBill.empsysnbr=client.clibillingatty" +
                    " left outer join employee cliResp on cliResp.empsysnbr=crtemployeeid" +
                    " where clisysnbr in (" + clients + ")";
            return reportSQL;
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

    }
}
