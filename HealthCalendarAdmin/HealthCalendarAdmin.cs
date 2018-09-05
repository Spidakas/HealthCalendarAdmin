
using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using HealthCalendarClasses;
using NLog;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace HealthCalendarAdmin
{
    public partial class HealthCalendarAdmin : Form
    {
        public HealthCalendarAdmin()
        {
            InitializeComponent();

            //Initialize NLog Targets and Rules
            var config = new NLog.Config.LoggingConfiguration();
            var logfile = new NLog.Targets.FileTarget("IdResult") { FileName = "BuildAddressBookEntryIdResult.txt" };
            config.AddRule(LogLevel.Debug, LogLevel.Fatal, logfile);
            NLog.LogManager.Configuration = config;

            bgWorker.DoWork += BgWorker_DoWork;
            bgWorker.ProgressChanged += BgWorker_ProgressChanged;
            //bgWorker.RunWorkerCompleted += BgWorker_RunWorkerCompleted;
        }

        HealthCalendarClasses.HealthCalendarClass c = new HealthCalendarClass();
        static string myconnstr = ConfigurationManager.ConnectionStrings["connstrng"].ConnectionString;

        private void HealthCalendarAdmin_Load(object sender, EventArgs e)
        {
            int rc;
            bool successA = c.GetTrustSettings(c);
            if (successA == false)
            {
                MessageBox.Show("Cannot connect to Health Calender database.");
            }

            //bool successB = c.GetGoogleClientSecret(c);
            //bool successC = c.GetGoogleAuthorization(c);
            bool successD = c.GetNHSNetAuthorization(c);
            if (successD == false)
            {
                MessageBox.Show("Unable to connect to NHS Net Server. Please contact your IT Department.");
            }
            bool successE = c.GetExchangeAuthorization(c);
            if (successE == false)
            {
                MessageBox.Show("Unable to connect to Exchange Server. Please contact your IT Department.");
            }


            // Create the ToolTip and associate with the Form container.
            ToolTip toolTip1 = new ToolTip();

            // Set up the delays for the ToolTip.
            toolTip1.AutoPopDelay = 10000;
            toolTip1.InitialDelay = 100;
            toolTip1.ReshowDelay = 500;
            // Force the ToolTip text to be displayed whether or not the form is active.
            toolTip1.ShowAlways = true;

            toolTip1.SetToolTip(btnGoogleEmailUpdate, "Update/Set a valid Google Email for the selected user");
            toolTip1.SetToolTip(btnGoogleEmailClear, "Remove the Google Email for the selected user");
            toolTip1.SetToolTip(btnGoogleCreateShare, "Create and share a Google calendar"); 
            toolTip1.SetToolTip(btnSampleGoogleData, "Set Google trial calendar data");
            toolTip1.SetToolTip(btnClearGoogleCalendar, "Clear Google calendar");
            toolTip1.SetToolTip(btnDeleteGoogleCalendar, "Delete Google calendar");

            toolTip1.SetToolTip(btnNHSNetEmailUpdate, "Update/Set a valid NHSNet Email for the selected user");
            toolTip1.SetToolTip(btnNHSNetEmailClear, "Remove the NHSNet Email for the selected user");
            toolTip1.SetToolTip(btnNHSNetCreateShare, "Create and share a NHSNet calendar");
            toolTip1.SetToolTip(btnSampleNHSNetData, "Set NHSNet trial calendar data");
            toolTip1.SetToolTip(btnClearNHSNetCalendar, "Clear NHSNet calendar");
            toolTip1.SetToolTip(btnDeleteNHSNetCalendar, "Delete NHSNet calendar");

            toolTip1.SetToolTip(btnExchangeEmailUpdate, "Update/Set a valid Exchange Email for the selected user");
            toolTip1.SetToolTip(btnExchangeEmailClear, "Remove the Exchange Email for the selected user");
            toolTip1.SetToolTip(btnExchangeCreateShare, "Create and share an Exchange calendar");
            toolTip1.SetToolTip(btnSampleExchangeData, "Set Exchange trial calendar data");
            toolTip1.SetToolTip(btnClearExchangeCalendar, "Clear Exchange calendar");
            toolTip1.SetToolTip(btnDeleteExchangeCalendar, "Delete Exchange calendar");

            dgvSubscribers.AutoGenerateColumns = false;
            dgvSubscribers.ColumnCount = 28;
            dgvSubscribers.Columns[0].HeaderText = "Subscriber ID";
            dgvSubscribers.Columns[0].DataPropertyName = "SubscriberID";
            dgvSubscribers.Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable;
            dgvSubscribers.Columns[0].Visible = false;

            dgvSubscribers.Columns[1].HeaderText = "Subscriber OID";
            dgvSubscribers.Columns[1].DataPropertyName = "SubscriberOID";
            dgvSubscribers.Columns[1].SortMode = DataGridViewColumnSortMode.NotSortable;
            dgvSubscribers.Columns[1].MinimumWidth = 95;
            dgvSubscribers.Columns[1].Width = 100;
            dgvSubscribers.Columns[1].Visible = false;

            dgvSubscribers.Columns[2].HeaderText = "Initials";
            dgvSubscribers.Columns[2].DataPropertyName = "Initials";
            dgvSubscribers.Columns[2].SortMode = DataGridViewColumnSortMode.NotSortable;
            dgvSubscribers.Columns[2].MinimumWidth = 40;
            dgvSubscribers.Columns[2].Width = 40;

            dgvSubscribers.Columns[3].HeaderText = "Title";
            dgvSubscribers.Columns[3].DataPropertyName = "Title";
            dgvSubscribers.Columns[3].SortMode = DataGridViewColumnSortMode.NotSortable;
            dgvSubscribers.Columns[3].MinimumWidth = 40;
            dgvSubscribers.Columns[3].Width = 55;

            dgvSubscribers.Columns[4].HeaderText = "First Name";
            dgvSubscribers.Columns[4].DataPropertyName = "Firstname";
            dgvSubscribers.Columns[4].SortMode = DataGridViewColumnSortMode.NotSortable;
            dgvSubscribers.Columns[4].MinimumWidth = 85;
            dgvSubscribers.Columns[4].Width = 100;

            dgvSubscribers.Columns[5].HeaderText = "Last Name";
            dgvSubscribers.Columns[5].DataPropertyName = "Surname";
            dgvSubscribers.Columns[5].SortMode = DataGridViewColumnSortMode.NotSortable;
            dgvSubscribers.Columns[5].MinimumWidth = 85;
            dgvSubscribers.Columns[5].Width = 120;

            dgvSubscribers.Columns[6].HeaderText = "Sex";
            dgvSubscribers.Columns[6].DataPropertyName = "Sex";
            dgvSubscribers.Columns[6].SortMode = DataGridViewColumnSortMode.NotSortable;
            dgvSubscribers.Columns[6].MinimumWidth = 45;
            dgvSubscribers.Columns[6].Width = 45;

            dgvSubscribers.Columns[7].HeaderText = "Modified Date";
            dgvSubscribers.Columns[7].DataPropertyName = "ModifiedDTTM";
            dgvSubscribers.Columns[7].SortMode = DataGridViewColumnSortMode.NotSortable;
            dgvSubscribers.Columns[7].MinimumWidth = 85;
            dgvSubscribers.Columns[7].Width = 120;
            dgvSubscribers.Columns[7].Visible = false;

            dgvSubscribers.Columns[8].HeaderText = "End Date";
            dgvSubscribers.Columns[8].DataPropertyName = "EndDTTM";
            dgvSubscribers.Columns[8].SortMode = DataGridViewColumnSortMode.NotSortable;
            dgvSubscribers.Columns[8].MinimumWidth = 85;
            dgvSubscribers.Columns[8].Width = 120;
            dgvSubscribers.Columns[8].Visible = false;

            dgvSubscribers.Columns[9].HeaderText = "Occupation";
            dgvSubscribers.Columns[9].DataPropertyName = "Occupation";
            dgvSubscribers.Columns[9].SortMode = DataGridViewColumnSortMode.NotSortable;
            dgvSubscribers.Columns[9].MinimumWidth = 85;
            dgvSubscribers.Columns[9].Width = 120;

            dgvSubscribers.Columns[10].HeaderText = "Main Identifier";
            dgvSubscribers.Columns[10].DataPropertyName = "MainIdentifier";
            dgvSubscribers.Columns[10].SortMode = DataGridViewColumnSortMode.NotSortable;
            dgvSubscribers.Columns[10].MinimumWidth = 80;
            dgvSubscribers.Columns[10].Width = 80;

            dgvSubscribers.Columns[11].HeaderText = "Source OID";
            dgvSubscribers.Columns[11].DataPropertyName = "SourceOID";
            dgvSubscribers.Columns[11].SortMode = DataGridViewColumnSortMode.NotSortable;
            dgvSubscribers.Columns[11].MinimumWidth = 85;
            dgvSubscribers.Columns[11].Width = 120;
            dgvSubscribers.Columns[11].Visible = false;

            dgvSubscribers.Columns[12].HeaderText = "Source Type";
            dgvSubscribers.Columns[12].DataPropertyName = "SourceType";
            dgvSubscribers.Columns[12].SortMode = DataGridViewColumnSortMode.NotSortable;
            dgvSubscribers.Columns[12].MinimumWidth = 85;
            dgvSubscribers.Columns[12].Width = 120;
            dgvSubscribers.Columns[12].Visible = false;

            dgvSubscribers.Columns[13].HeaderText = "Owner Organisation OID";
            dgvSubscribers.Columns[13].DataPropertyName = "OwnerOrganisationOID";
            dgvSubscribers.Columns[13].SortMode = DataGridViewColumnSortMode.NotSortable;
            dgvSubscribers.Columns[13].MinimumWidth = 85;
            dgvSubscribers.Columns[13].Width = 120;
            dgvSubscribers.Columns[13].Visible = false;

            dgvSubscribers.Columns[14].HeaderText = "RONEOID";
            dgvSubscribers.Columns[14].DataPropertyName = "RONEOID";
            dgvSubscribers.Columns[14].SortMode = DataGridViewColumnSortMode.NotSortable;
            dgvSubscribers.Columns[14].MinimumWidth = 85;
            dgvSubscribers.Columns[14].Width = 120;
            dgvSubscribers.Columns[14].Visible = false;

            dgvSubscribers.Columns[15].HeaderText = "UITYPCODE";
            dgvSubscribers.Columns[15].DataPropertyName = "UITYPCODE";
            dgvSubscribers.Columns[15].SortMode = DataGridViewColumnSortMode.NotSortable;
            dgvSubscribers.Columns[15].MinimumWidth = 85;
            dgvSubscribers.Columns[15].Width = 120;
            dgvSubscribers.Columns[15].Visible = false;

            dgvSubscribers.Columns[16].HeaderText = "Google Email";
            dgvSubscribers.Columns[16].DataPropertyName = "GoogleEmail";
            dgvSubscribers.Columns[16].SortMode = DataGridViewColumnSortMode.NotSortable;
            dgvSubscribers.Columns[16].MinimumWidth = 95;
            dgvSubscribers.Columns[16].Width = 180;

            dgvSubscribers.Columns[17].HeaderText = "Google Calendar ID";
            dgvSubscribers.Columns[17].DataPropertyName = "GoogleCalendarID";
            dgvSubscribers.Columns[17].SortMode = DataGridViewColumnSortMode.NotSortable;
            dgvSubscribers.Columns[17].MinimumWidth = 125;
            dgvSubscribers.Columns[17].Width = 180;

            dgvSubscribers.Columns[18].HeaderText = "Outlook Email";
            dgvSubscribers.Columns[18].DataPropertyName = "OutlookEmail";
            dgvSubscribers.Columns[18].SortMode = DataGridViewColumnSortMode.NotSortable;
            dgvSubscribers.Columns[18].MinimumWidth = 85;
            dgvSubscribers.Columns[18].Width = 120;
            dgvSubscribers.Columns[18].Visible = false;

            dgvSubscribers.Columns[19].HeaderText = "Outlook Calendar ID";
            dgvSubscribers.Columns[19].DataPropertyName = "OutlookCalendarID";
            dgvSubscribers.Columns[19].SortMode = DataGridViewColumnSortMode.NotSortable;
            dgvSubscribers.Columns[19].MinimumWidth = 85;
            dgvSubscribers.Columns[19].Width = 120;
            dgvSubscribers.Columns[19].Visible = false;

            dgvSubscribers.Columns[20].HeaderText = "Exchange Email";
            dgvSubscribers.Columns[20].DataPropertyName = "ExchangeEmail";
            dgvSubscribers.Columns[20].SortMode = DataGridViewColumnSortMode.NotSortable;
            dgvSubscribers.Columns[20].MinimumWidth = 95;
            dgvSubscribers.Columns[20].Width = 180;
            //dgvSubscribers.Columns[20].Visible = false;

            dgvSubscribers.Columns[21].HeaderText = "Exchange Calendar ID";
            dgvSubscribers.Columns[21].DataPropertyName = "ExchangeCalendarID";
            dgvSubscribers.Columns[21].SortMode = DataGridViewColumnSortMode.NotSortable;
            dgvSubscribers.Columns[21].MinimumWidth = 125;
            dgvSubscribers.Columns[21].Width = 180;
            //dgvSubscribers.Columns[21].Visible = false;

            dgvSubscribers.Columns[22].HeaderText = "Exchange Calendar Name";
            dgvSubscribers.Columns[22].DataPropertyName = "ExchangeCalendarName";
            dgvSubscribers.Columns[22].SortMode = DataGridViewColumnSortMode.NotSortable;
            dgvSubscribers.Columns[22].MinimumWidth = 125;
            //dgvSubscribers.Columns[22].Width = 180;

            dgvSubscribers.Columns[23].HeaderText = "NHSNet Email";
            dgvSubscribers.Columns[23].DataPropertyName = "NHSNetEmail";
            dgvSubscribers.Columns[23].SortMode = DataGridViewColumnSortMode.NotSortable;
            dgvSubscribers.Columns[23].MinimumWidth = 95;
            dgvSubscribers.Columns[23].Width = 180;

            dgvSubscribers.Columns[24].HeaderText = "NHSNet Calendar ID";
            dgvSubscribers.Columns[24].DataPropertyName = "NHSNetCalendarID";
            dgvSubscribers.Columns[24].SortMode = DataGridViewColumnSortMode.NotSortable;
            dgvSubscribers.Columns[24].MinimumWidth = 125;
            dgvSubscribers.Columns[24].Width = 180;

            dgvSubscribers.Columns[25].HeaderText = "NHSNet Calendar Name";
            dgvSubscribers.Columns[25].DataPropertyName = "NHSNetCalendarName";
            dgvSubscribers.Columns[25].SortMode = DataGridViewColumnSortMode.NotSortable;
            dgvSubscribers.Columns[25].MinimumWidth = 125;
            dgvSubscribers.Columns[25].Width = 180;

            dgvSubscribers.Columns[26].HeaderText = "Apple Email";
            dgvSubscribers.Columns[26].DataPropertyName = "AppleEmail";
            dgvSubscribers.Columns[26].SortMode = DataGridViewColumnSortMode.NotSortable;
            dgvSubscribers.Columns[26].MinimumWidth = 95;
            dgvSubscribers.Columns[26].Width = 180;
            dgvSubscribers.Columns[26].Visible = false;

            dgvSubscribers.Columns[27].HeaderText = "Apple Calendar ID";
            dgvSubscribers.Columns[27].DataPropertyName = "AppleCalendarID";
            dgvSubscribers.Columns[27].SortMode = DataGridViewColumnSortMode.NotSortable;
            dgvSubscribers.Columns[27].MinimumWidth = 125;
            dgvSubscribers.Columns[27].Width = 180;
            dgvSubscribers.Columns[27].Visible = false;
    
            searchSubscribers(txtboxSearchFirstname.Text, txtboxSearchLastname.Text, comboBoxSex.Text, comboBoxTitle.Text, comboBoxOccupation.Text, txtboxSearchMainIdentifier.Text);
            rc = dgvSubscribers.CurrentCell.RowIndex;
            dgvSubscribers.CurrentCell = dgvSubscribers.Rows[rc].Cells[2]; //First visible column
            dgvSubscribers.Rows[rc].Selected = true;

        }

        private void HealthCalendarAdmin_Shown(object sender, EventArgs e)
        {
            txtboxSearchFirstname.Focus();
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            c.GoogleEmail = txtboxGoogleEmail.Text;
            if (c.IsValidEmail(c.GoogleEmail)) {
                bool success = c.UpdateGoogle(c);
                if (success == true)
                {
                    MessageBox.Show("Google subscriber has been successfuly updated.");
                }
                else
                {
                    MessageBox.Show("Failed to update Google subscriber. Please try again.");
                }

            } else {
                MessageBox.Show("Invalid Google email address. Please try again.");
            }
            searchSubscribersAndSelectRow(txtboxSearchFirstname.Text, txtboxSearchLastname.Text, comboBoxSex.Text, comboBoxTitle.Text, comboBoxOccupation.Text, txtboxSearchMainIdentifier.Text);
        }

        private void btnNHSNetEmailUpdate_Click(object sender, EventArgs e)
        {
            c.NHSNetEmail = txtboxNHSNetEmail.Text;
            if (c.IsValidEmail(c.NHSNetEmail))
            {
                bool success = c.UpdateNHSNet(c);
                if (success == true)
                {
                    MessageBox.Show("NHSNet subscriber has been successfuly updated.");
                }
                else
                {
                    MessageBox.Show("Failed to update NHSNet subscriber. Please try again.");
                }

            }
            else
            {
                MessageBox.Show("Invalid NHSNet email address. Please try again.");
            }
            searchSubscribersAndSelectRow(txtboxSearchFirstname.Text, txtboxSearchLastname.Text, comboBoxSex.Text, comboBoxTitle.Text, comboBoxOccupation.Text, txtboxSearchMainIdentifier.Text);
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(c.GoogleCalendarID))
            {
                c.GoogleEmail = "";
                bool success = c.UpdateGoogle(c);
                if (success == true)
                {
                    MessageBox.Show("Google subscribers email has been successfuly removed.");        
                }
                else
                {
                    MessageBox.Show("Failed to remove subscribers Google email. Please try again.");
                }
                searchSubscribersAndSelectRow(txtboxSearchFirstname.Text, txtboxSearchLastname.Text, comboBoxSex.Text, comboBoxTitle.Text, comboBoxOccupation.Text, txtboxSearchMainIdentifier.Text);
            }                
        }

        private void dgvSubscribers_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            int rowIndex = e.RowIndex;
            txtboxGoogleEmail.Text = dgvSubscribers.Rows[rowIndex].Cells[16].Value.ToString();
        }

        private void HealthCalendarAdmin_Leave(object sender, EventArgs e)
        {
            this.Close();
        }

        public void Clear()
        {
            txtboxGoogleEmail.Text = "";
            txtboxNHSNetEmail.Text = "";
            txtboxExchangeEmail.Text = "";

        }


        private void btnGoogleCreateShare_Click(object sender, EventArgs e)
        {
            bool isSuccess = false;
            c.GoogleEmail = txtboxGoogleEmail.Text;
            if (c.GoogleCalendarID != null)
            {
                progressBar.Maximum = 100;
                //lblStatus.ForeColor = Color.Red;
                //lblStatus.Text = "Counting...";
                bgWorker.WorkerReportsProgress = true;
                bgWorker.RunWorkerAsync();
                isSuccess = c.CreateShareGoogleDiary(c);
                if (isSuccess == true)
                {
                    MessageBox.Show("The Google Calendar has been successfuly created and shared.");
                }
                else
                {
                    MessageBox.Show("Failed to create Google Calendar. Try Again .");
                }
            }
            searchSubscribersAndSelectRow(txtboxSearchFirstname.Text, txtboxSearchLastname.Text, comboBoxSex.Text, comboBoxTitle.Text, comboBoxOccupation.Text, txtboxSearchMainIdentifier.Text);
            Clear();
        }

        private void btnNHSNetCreateShare_Click(object sender, EventArgs e)
        {
            bool isSuccess = false;
            c.NHSNetEmail = txtboxNHSNetEmail.Text;
            if (c.NHSNetCalendarID != null)
            {
                progressBar.Maximum = 100;
                //lblStatus.ForeColor = Color.Red;
                //lblStatus.Text = "Counting...";
                bgWorker.WorkerReportsProgress = true;
                bgWorker.RunWorkerAsync();
                isSuccess = c.CreateShareNHSNetDiary(c);
                if (isSuccess == true)
                {
                    MessageBox.Show("The NHSNet Calendar has been successfuly created and shared.");
                }
                else
                {
                    MessageBox.Show("Failed to create NHSNet Calendar. Try Again .");
                }
            }
            searchSubscribersAndSelectRow(txtboxSearchFirstname.Text, txtboxSearchLastname.Text, comboBoxSex.Text, comboBoxTitle.Text, comboBoxOccupation.Text, txtboxSearchMainIdentifier.Text);
            Clear();
        }

        private void btnSampleGoogleData_Click(object sender, EventArgs e)
        {
            bool isSuccess = false;
            c.GoogleEmail = txtboxGoogleEmail.Text;
            if (c.GoogleCalendarID != null)
            {
                //progressBar.Maximum = 100;
                //bgWorker.WorkerReportsProgress = true;
                
                isSuccess = c.CreateSampleGoogleCalendarData(c);
                //bgWorker.RunWorkerAsync();
                if (isSuccess == true)
                {
                    MessageBox.Show("The Sample Google Calendar Data has been successfuly created.");
                }
                else
                {
                    MessageBox.Show("Failed to create Sample Google Calendar Data. Try Again .");
                }
            }
            searchSubscribersAndSelectRow(txtboxSearchFirstname.Text, txtboxSearchLastname.Text, comboBoxSex.Text, comboBoxTitle.Text, comboBoxOccupation.Text, txtboxSearchMainIdentifier.Text);
        }

        private void btnSampleNHSNetData_Click(object sender, EventArgs e)
        {
            bool isSuccess = false;
            c.NHSNetEmail = txtboxNHSNetEmail.Text;
            if (c.NHSNetCalendarID != null)
            {
                //progressBar.Maximum = 100;
                //bgWorker.WorkerReportsProgress = true;

                isSuccess = c.CreateSampleNHSNetCalendarData(c);
                //bgWorker.RunWorkerAsync();
                if (isSuccess == true)
                {
                    MessageBox.Show("The Sample NHSNet Calendar Data has been successfuly created.");
                }
                else
                {
                    MessageBox.Show("Failed to create Sample NHSNet Calendar Data. Try Again .");
                }
            }
            searchSubscribersAndSelectRow(txtboxSearchFirstname.Text, txtboxSearchLastname.Text, comboBoxSex.Text, comboBoxTitle.Text, comboBoxOccupation.Text, txtboxSearchMainIdentifier.Text);
        }

        private void BgWorker_DoWork(object sender, DoWorkEventArgs e)
        {

            for (var Counter = 1; Counter <= progressBar.Maximum; Counter++)
            {
                bgWorker.ReportProgress(Counter);
                System.Threading.Thread.Sleep(50);
            }
        }

        private void BgWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            lblPercent.Text = e.ProgressPercentage.ToString();
            progressBar.Value = e.ProgressPercentage;
        }

        private void BgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            //lblStatus.ForeColor = Color.Green;
            //lblStatus.Text = "Done";
        }

        private void btnDeleteGoogleCalendar_Click(object sender, EventArgs e)
        {
            bool isSuccess = false;

            if (c.GoogleCalendarID != null)
            {
   

                isSuccess = c.DeleteSecondaryGoogleCalendar(c);
                if (isSuccess == true)
                {
                    MessageBox.Show("The selected Google Calendar has been successfuly deleted.");
                }
                else
                {
                    MessageBox.Show("Failed to delete Google Calendar. Try Again .");
                }
            }
            searchSubscribersAndSelectRow(txtboxSearchFirstname.Text, txtboxSearchLastname.Text, comboBoxSex.Text, comboBoxTitle.Text, comboBoxOccupation.Text, txtboxSearchMainIdentifier.Text);
        }

        private void btnDeleteNHSNetCalendar_Click(object sender, EventArgs e)
        {
            bool isSuccess = false;

            if (c.NHSNetCalendarID != null)
            {


                isSuccess = c.DeleteNHSNetCalendar(c);
                if (isSuccess == true)
                {
                    MessageBox.Show("The selected Google Calendar has been successfuly deleted.");
                }
                else
                {
                    MessageBox.Show("Failed to delete Google Calendar. Try Again .");
                }
            }            
            searchSubscribersAndSelectRow(txtboxSearchFirstname.Text, txtboxSearchLastname.Text, comboBoxSex.Text, comboBoxTitle.Text, comboBoxOccupation.Text, txtboxSearchMainIdentifier.Text);
        }

        private void btnClearGoogleCalendar_Click(object sender, EventArgs e)
        {
            bool isSuccess = false;

            c.GoogleEmail = txtboxGoogleEmail.Text;
            if (c.GoogleCalendarID != null)
            {


                isSuccess = c.DeleteGoogleCalendarEvents(c);
                if (isSuccess == true)
                {
                    MessageBox.Show("The Google Calendar has been successfuly cleared.");
                }
                else
                {
                    MessageBox.Show("Failed to clear the Google Calendar. Try Again .");
                }
            }
            searchSubscribersAndSelectRow(txtboxSearchFirstname.Text, txtboxSearchLastname.Text, comboBoxSex.Text, comboBoxTitle.Text, comboBoxOccupation.Text, txtboxSearchMainIdentifier.Text);
        }

        private void btnClearNHSNetCalendar_Click(object sender, EventArgs e)
        {
            bool isSuccess = false;

            c.NHSNetEmail = txtboxNHSNetEmail.Text;
            if (c.NHSNetCalendarID != null)
            {


                isSuccess = c.DeleteNHSNetCalendarEvents(c);
                if (isSuccess == true)
                {
                    MessageBox.Show("The NHSNet Calendar has been successfuly cleared.");
                }
                else
                {
                    MessageBox.Show("Failed to clear NHSNet Calendar. Try Again .");
                }
            }            
            searchSubscribersAndSelectRow(txtboxSearchFirstname.Text, txtboxSearchLastname.Text, comboBoxSex.Text, comboBoxTitle.Text, comboBoxOccupation.Text, txtboxSearchMainIdentifier.Text);
        }

        private void btnNHSNetEmailClear_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(c.NHSNetCalendarID))
            {
                c.NHSNetEmail = "";
                bool success = c.UpdateNHSNet(c);
                if (success == true)
                {
                    MessageBox.Show("NHSNet subscribers email has been successfuly removed.");
                }
                else
                {
                    MessageBox.Show("Failed to remove subscribers NHS email. Please try again.");
                }
                searchSubscribersAndSelectRow(txtboxSearchFirstname.Text, txtboxSearchLastname.Text, comboBoxSex.Text, comboBoxTitle.Text, comboBoxOccupation.Text, txtboxSearchMainIdentifier.Text);
            }
        }

        private void btnSendExchangeEmail_Click(object sender, EventArgs e)
        {
           
        }

        private void dgvSubscribers_RowEnter(object sender, DataGridViewCellEventArgs e)
        {
            int rowIndex = e.RowIndex;
            c.SubscriberID = dgvSubscribers.Rows[rowIndex].Cells[0].Value.ToString();
            c.SubscriberOID = dgvSubscribers.Rows[rowIndex].Cells[1].Value.ToString();
            c.Title = dgvSubscribers.Rows[rowIndex].Cells[3].Value.ToString();
            c.FirstName = dgvSubscribers.Rows[rowIndex].Cells[4].Value.ToString();
            c.LastName = dgvSubscribers.Rows[rowIndex].Cells[5].Value.ToString();
            txtboxGoogleEmail.Text = dgvSubscribers.Rows[rowIndex].Cells[16].Value.ToString();
            c.GoogleEmail = dgvSubscribers.Rows[rowIndex].Cells[16].Value.ToString();
            c.GoogleCalendarID = dgvSubscribers.Rows[rowIndex].Cells[17].Value.ToString();

            txtboxExchangeEmail.Text = dgvSubscribers.Rows[rowIndex].Cells[20].Value.ToString();
            c.ExchangeEmail = dgvSubscribers.Rows[rowIndex].Cells[20].Value.ToString();
            c.ExchangeCalendarID = dgvSubscribers.Rows[rowIndex].Cells[21].Value.ToString();
            c.ExchangeCalendarName = dgvSubscribers.Rows[rowIndex].Cells[22].Value.ToString();

            txtboxNHSNetEmail.Text = dgvSubscribers.Rows[rowIndex].Cells[23].Value.ToString();
            c.NHSNetEmail = dgvSubscribers.Rows[rowIndex].Cells[23].Value.ToString();
            c.NHSNetCalendarID = dgvSubscribers.Rows[rowIndex].Cells[24].Value.ToString();
            c.NHSNetCalendarName = dgvSubscribers.Rows[rowIndex].Cells[25].Value.ToString();

            //If Valid Email and no CalenderID enabled/visible G+ / hide 3
            //If Valid CalenderID then hide/disable g+ = enable other 3
            if (c.IsValidEmail(c.GoogleEmail))
            {
                if (string.IsNullOrEmpty(c.GoogleCalendarID))                
                {
                    lblGoogleEmail.Visible = true;
                    lblGoogleEmail.Enabled = true;
                    txtboxGoogleEmail.Visible = true;
                    txtboxGoogleEmail.Enabled = true;

                    btnGoogleEmailUpdate.Visible = true;
                    btnGoogleEmailUpdate.Enabled = true;
                    btnGoogleEmailClear.Visible = true;
                    btnGoogleEmailClear.Enabled = true;
                    btnGoogleCreateShare.Visible = true;
                    btnGoogleCreateShare.Enabled = true;

                    btnSampleGoogleData.Visible = false;
                    btnSampleGoogleData.Enabled = false;
                    btnClearGoogleCalendar.Visible = false;
                    btnClearGoogleCalendar.Enabled = false;
                    btnDeleteGoogleCalendar.Visible = false;
                    btnDeleteGoogleCalendar.Enabled = false;
                }
                else
                {
                    lblGoogleEmail.Visible = false;
                    lblGoogleEmail.Enabled = false;
                    txtboxGoogleEmail.Visible = false;
                    txtboxGoogleEmail.Enabled = false;
                    btnGoogleEmailUpdate.Visible = false;
                    btnGoogleEmailUpdate.Enabled = false;
                    btnGoogleEmailClear.Visible = false;
                    btnGoogleEmailClear.Enabled = false;
                    btnGoogleCreateShare.Visible = false;
                    btnGoogleCreateShare.Enabled = false;

                    btnSampleGoogleData.Visible = true;
                    btnSampleGoogleData.Enabled = true;
                    btnClearGoogleCalendar.Visible = true;
                    btnClearGoogleCalendar.Enabled = true;
                    btnDeleteGoogleCalendar.Visible = true;
                    btnDeleteGoogleCalendar.Enabled = true;
                }
            }
            else
            {
                lblGoogleEmail.Visible = true;
                lblGoogleEmail.Enabled = true;
                txtboxGoogleEmail.Visible = true;
                txtboxGoogleEmail.Enabled = true;

                btnGoogleEmailUpdate.Visible = true;
                btnGoogleEmailUpdate.Enabled = true;
                btnGoogleEmailClear.Visible = true;
                btnGoogleEmailClear.Enabled = true;

                btnGoogleCreateShare.Visible = false;
                btnGoogleCreateShare.Enabled = false;

                btnSampleGoogleData.Visible = false;
                btnSampleGoogleData.Enabled = false;
                btnClearGoogleCalendar.Visible = false;
                btnClearGoogleCalendar.Enabled = false;
                btnDeleteGoogleCalendar.Visible = false;
                btnDeleteGoogleCalendar.Enabled = false;
            }

            //NHSNet
            if (c.IsValidEmail(c.NHSNetEmail))
            {
                if (string.IsNullOrEmpty(c.NHSNetCalendarID))
                {
                    lblNHSNetEmail.Visible = true;
                    lblNHSNetEmail.Enabled = true;
                    txtboxNHSNetEmail.Visible = true;
                    txtboxNHSNetEmail.Enabled = true;

                    btnNHSNetEmailUpdate.Visible = true;
                    btnNHSNetEmailUpdate.Enabled = true;
                    btnNHSNetEmailClear.Visible = true;
                    btnNHSNetEmailClear.Enabled = true;
                    btnNHSNetCreateShare.Visible = true;
                    btnNHSNetCreateShare.Enabled = true;

                    btnSampleNHSNetData.Visible = false;
                    btnSampleNHSNetData.Enabled = false;
                    btnClearNHSNetCalendar.Visible = false;
                    btnClearNHSNetCalendar.Enabled = false;
                    btnDeleteNHSNetCalendar.Visible = false;
                    btnDeleteNHSNetCalendar.Enabled = false;
                }
                else
                {
                    lblNHSNetEmail.Visible = false;
                    lblNHSNetEmail.Enabled = false;
                    txtboxNHSNetEmail.Visible = false;
                    txtboxNHSNetEmail.Enabled = false;
                    btnNHSNetEmailUpdate.Visible = false;
                    btnNHSNetEmailUpdate.Enabled = false;
                    btnNHSNetEmailClear.Visible = false;
                    btnNHSNetEmailClear.Enabled = false;
                    btnNHSNetCreateShare.Visible = false;
                    btnNHSNetCreateShare.Enabled = false;

                    btnSampleNHSNetData.Visible = true;
                    btnSampleNHSNetData.Enabled = true;
                    btnClearNHSNetCalendar.Visible = true;
                    btnClearNHSNetCalendar.Enabled = true;
                    btnDeleteNHSNetCalendar.Visible = true;
                    btnDeleteNHSNetCalendar.Enabled = true;
                }
            }
            else
            {
                lblNHSNetEmail.Visible = true;
                lblNHSNetEmail.Enabled = true;
                txtboxNHSNetEmail.Visible = true;
                txtboxNHSNetEmail.Enabled = true;

                btnNHSNetEmailUpdate.Visible = true;
                btnNHSNetEmailUpdate.Enabled = true;
                btnNHSNetEmailClear.Visible = true;
                btnNHSNetEmailClear.Enabled = true;

                btnNHSNetCreateShare.Visible = false;
                btnNHSNetCreateShare.Enabled = false;

                btnSampleNHSNetData.Visible = false;
                btnSampleNHSNetData.Enabled = false;
                btnClearNHSNetCalendar.Visible = false;
                btnClearNHSNetCalendar.Enabled = false;
                btnDeleteNHSNetCalendar.Visible = false;
                btnDeleteNHSNetCalendar.Enabled = false;
            }

            //Exchange
            if (c.IsValidEmail(c.ExchangeEmail))
            {
                if (string.IsNullOrEmpty(c.ExchangeCalendarID))
                {
                    lblExchangeEmail.Visible = true;
                    lblExchangeEmail.Enabled = true;
                    txtboxExchangeEmail.Visible = true;
                    txtboxExchangeEmail.Enabled = true;

                    btnExchangeEmailUpdate.Visible = true;
                    btnExchangeEmailUpdate.Enabled = true;
                    btnExchangeEmailClear.Visible = true;
                    btnExchangeEmailClear.Enabled = true;
                    btnExchangeCreateShare.Visible = true;
                    btnExchangeCreateShare.Enabled = true;

                    btnSampleExchangeData.Visible = false;
                    btnSampleExchangeData.Enabled = false;
                    btnClearExchangeCalendar.Visible = false;
                    btnClearExchangeCalendar.Enabled = false;
                    btnDeleteExchangeCalendar.Visible = false;
                    btnDeleteExchangeCalendar.Enabled = false;
                }
                else
                {
                    lblExchangeEmail.Visible = false;
                    lblExchangeEmail.Enabled = false;
                    txtboxExchangeEmail.Visible = false;
                    txtboxExchangeEmail.Enabled = false;
                    btnExchangeEmailUpdate.Visible = false;
                    btnExchangeEmailUpdate.Enabled = false;
                    btnExchangeEmailClear.Visible = false;
                    btnExchangeEmailClear.Enabled = false;
                    btnExchangeCreateShare.Visible = false;
                    btnExchangeCreateShare.Enabled = false;

                    btnSampleExchangeData.Visible = true;
                    btnSampleExchangeData.Enabled = true;
                    btnClearExchangeCalendar.Visible = true;
                    btnClearExchangeCalendar.Enabled = true;
                    btnDeleteExchangeCalendar.Visible = true;
                    btnDeleteExchangeCalendar.Enabled = true;
                }
            }
            else
            {
                lblExchangeEmail.Visible = true;
                lblExchangeEmail.Enabled = true;
                txtboxExchangeEmail.Visible = true;
                txtboxExchangeEmail.Enabled = true;

                btnExchangeEmailUpdate.Visible = true;
                btnExchangeEmailUpdate.Enabled = true;
                btnExchangeEmailClear.Visible = true;
                btnExchangeEmailClear.Enabled = true;

                btnExchangeCreateShare.Visible = false;
                btnExchangeCreateShare.Enabled = false;

                btnSampleExchangeData.Visible = false;
                btnSampleExchangeData.Enabled = false;
                btnClearExchangeCalendar.Visible = false;
                btnClearExchangeCalendar.Enabled = false;
                btnDeleteExchangeCalendar.Visible = false;
                btnDeleteExchangeCalendar.Enabled = false;
            }


        }

        private void dgvSubscribers_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            searchSubscribers(txtboxSearchFirstname.Text, txtboxSearchLastname.Text, comboBoxSex.Text, comboBoxTitle.Text, comboBoxOccupation.Text, txtboxSearchMainIdentifier.Text);

        }
        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            searchSubscribers(txtboxSearchFirstname.Text, txtboxSearchLastname.Text, comboBoxSex.Text, comboBoxTitle.Text, comboBoxOccupation.Text, txtboxSearchMainIdentifier.Text);
        }

        private void comboBoxOccupation_SelectedIndexChanged(object sender, EventArgs e)
        {
            searchSubscribers(txtboxSearchFirstname.Text, txtboxSearchLastname.Text, comboBoxSex.Text, comboBoxTitle.Text, comboBoxOccupation.Text, txtboxSearchMainIdentifier.Text);
        }

        private void comboBoxSex_SelectedIndexChanged(object sender, EventArgs e)
        {
            searchSubscribers(txtboxSearchFirstname.Text, txtboxSearchLastname.Text, comboBoxSex.Text, comboBoxTitle.Text, comboBoxOccupation.Text, txtboxSearchMainIdentifier.Text);
        }

        private void comboBoxTitle_SelectedIndexChanged(object sender, EventArgs e)
        {
            searchSubscribers(txtboxSearchFirstname.Text, txtboxSearchLastname.Text, comboBoxSex.Text, comboBoxTitle.Text, comboBoxOccupation.Text, txtboxSearchMainIdentifier.Text);
        }

        private void txtboxSearchMainIdentifier_TextChanged(object sender, EventArgs e)
        {
            searchSubscribers(txtboxSearchFirstname.Text, txtboxSearchLastname.Text, comboBoxSex.Text, comboBoxTitle.Text, comboBoxOccupation.Text, txtboxSearchMainIdentifier.Text);
        }



        public void searchSubscribersAndSelectRow(string keyword1, string keyword2, string keyword3, string keyword4, string keyword5, string keyword6)
        {
            int rc;

            rc = dgvSubscribers.CurrentCell.RowIndex;
            searchSubscribers(txtboxSearchFirstname.Text, txtboxSearchLastname.Text, comboBoxSex.Text, comboBoxTitle.Text, comboBoxOccupation.Text, txtboxSearchMainIdentifier.Text);
            dgvSubscribers.CurrentCell = dgvSubscribers.Rows[rc].Cells[2]; //First visible column
            dgvSubscribers.Rows[rc].Selected = true;
        }



        public void searchSubscribers(string keyword1, string keyword2, string keyword3, string keyword4, string keyword5, string keyword6)
        {
            SqlConnection conn = new SqlConnection(myconnstr);
            //SqlDataAdapter sda = new SqlDataAdapter("SELECT * FROM Subscribers WHERE Firstname LIKE '%" + keyword1 + "%' AND Surname LIKE '%" + keyword2 + "%' AND Sex LIKE '%" + keyword3 + "%' AND MainIdentifier LIKE '%" + keyword6 + "%'", conn);

            SqlDataAdapter sda = new SqlDataAdapter("SELECT * FROM Subscribers WHERE IsNull(Firstname,'') LIKE '%" + keyword1 + "%' AND IsNull(Surname,'') LIKE '%" + keyword2 + "%' AND IsNull(Sex,'') LIKE '%" + keyword3 + "%' AND IsNull(Title,'') LIKE '" + keyword4 + "%' AND IsNull(Occupation,'') LIKE '%" + keyword5 + "%' AND IsNull(MainIdentifier,'') LIKE '%" + keyword6 + "%'", conn);
            DataTable dt = new DataTable();
            sda.Fill(dt);
            dgvSubscribers.DataSource = dt;
        }



        private void btnClearSearch_Click(object sender, EventArgs e)
        {
            txtboxSearchFirstname.Text = "";
            txtboxSearchLastname.Text = ""; ;
            comboBoxSex.Text="";
            comboBoxTitle.Text="";
            comboBoxOccupation.Text="";
            txtboxSearchMainIdentifier.Text="";
            searchSubscribers(txtboxSearchFirstname.Text, txtboxSearchLastname.Text, comboBoxSex.Text, comboBoxTitle.Text, comboBoxOccupation.Text, txtboxSearchMainIdentifier.Text);
        }



        private void btnAdd_Click(object sender, EventArgs e)
        {
            //c.Title = txtboxTitle.Text;
            //c.FirstName = txtboxFirstName.Text;
            //c.LastName = txtboxLastName.Text;
            //c.SubscriberID = txtboxSubscriberID.Text;
            //c.GoogleEmail = txtboxGoogleEmail.Text;
            //c.GoogleCalendarID = txtboxGoogleCalendarID.Text;

            //bool success = c.Insert(c);
            //if (success == true)
            //{
            //    MessageBox.Show("New Subscriber Successfuly Inserted");
            //    Clear();
            //}
            //else
            //{
            //    MessageBox.Show("Failed to add New Subscriber. Try Again.");
            //}
            //DataTable dt = c.Select();
            //dgvSubscribers.DataSource = dt;
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            //c.SubscriberID = txtboxSubscriberID.Text;
            //bool success = c.Delete(c);
            //if (success == true)
            //{
            //    MessageBox.Show("Subscriber has been successfuly Deleted.");
            //    DataTable dt = c.Select();
            //    dgvSubscribers.DataSource = dt;
            //    Clear();
            //}
            //else
            //{
            //    MessageBox.Show("Failed to Delete Subscriber. Try Again .");
            //}
        }

        private void txtboxGoogleEmail_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtboxGoogleCalendarID_TextChanged(object sender, EventArgs e)
        {

        }

        private void btnExchangeEmailUpdate_Click(object sender, EventArgs e)
        {
            c.ExchangeEmail = txtboxExchangeEmail.Text;
            if (c.IsValidEmail(c.ExchangeEmail))
            {
                bool success = c.UpdateExchange(c);
                if (success == true)
                {
                    MessageBox.Show("Exchange subscriber has been successfuly updated.");
                }
                else
                {
                    MessageBox.Show("Failed to update Exchange subscriber. Please try again.");
                }

            }
            else
            {
                MessageBox.Show("Invalid Exchange email address. Please try again.");
            }
            searchSubscribersAndSelectRow(txtboxSearchFirstname.Text, txtboxSearchLastname.Text, comboBoxSex.Text, comboBoxTitle.Text, comboBoxOccupation.Text, txtboxSearchMainIdentifier.Text);
        }

        private void btnExchangeEmailClear_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(c.ExchangeCalendarID))
            {
                c.ExchangeEmail = "";
                bool success = c.UpdateExchange(c);
                if (success == true)
                {
                    MessageBox.Show("Exchange subscribers email has been successfuly removed.");
                }
                else
                {
                    MessageBox.Show("Failed to remove subscribers Exchange email. Please try again.");
                }
                searchSubscribersAndSelectRow(txtboxSearchFirstname.Text, txtboxSearchLastname.Text, comboBoxSex.Text, comboBoxTitle.Text, comboBoxOccupation.Text, txtboxSearchMainIdentifier.Text);
            }
        }

        private void btnExchangeCreateShare_Click(object sender, EventArgs e)
        {
            bool isSuccess = false;
            c.ExchangeEmail = txtboxExchangeEmail.Text;
            if (c.ExchangeCalendarID != null)
            {
                progressBar.Maximum = 100;
                //lblStatus.ForeColor = Color.Red;
                //lblStatus.Text = "Counting...";
                bgWorker.WorkerReportsProgress = true;
                bgWorker.RunWorkerAsync();
                isSuccess = c.CreateShareExchangeDiary(c);
                if (isSuccess == true)
                {
                    MessageBox.Show("The Exchange Calendar has been successfuly created and shared.");
                }
                else
                {
                    MessageBox.Show("Failed to create Exchange Calendar. Try Again .");
                }
            }
            searchSubscribersAndSelectRow(txtboxSearchFirstname.Text, txtboxSearchLastname.Text, comboBoxSex.Text, comboBoxTitle.Text, comboBoxOccupation.Text, txtboxSearchMainIdentifier.Text);
            Clear();
        }

        private void btnSampleExchangeData_Click(object sender, EventArgs e)
        {
            bool isSuccess = false;
            c.ExchangeEmail = txtboxExchangeEmail.Text;
            if (c.ExchangeCalendarID != null)
            {
                //progressBar.Maximum = 100;
                //bgWorker.WorkerReportsProgress = true;

                isSuccess = c.CreateSampleExchangeCalendarData(c);
                //bgWorker.RunWorkerAsync();
                if (isSuccess == true)
                {
                    MessageBox.Show("The Sample Exchange Calendar Data has been successfuly created.");
                }
                else
                {
                    MessageBox.Show("Failed to create Sample Exchange Calendar Data. Try Again .");
                }
            }
            searchSubscribersAndSelectRow(txtboxSearchFirstname.Text, txtboxSearchLastname.Text, comboBoxSex.Text, comboBoxTitle.Text, comboBoxOccupation.Text, txtboxSearchMainIdentifier.Text);
        }

        private void btnClearExchangeCalendar_Click(object sender, EventArgs e)
        {
            bool isSuccess = false;

            c.ExchangeEmail = txtboxExchangeEmail.Text;
            if (c.ExchangeCalendarID != null)
            {


                isSuccess = c.DeleteExchangeCalendarEvents(c);
                if (isSuccess == true)
                {
                    MessageBox.Show("The Exchange Calendar has been successfuly cleared.");
                }
                else
                {
                    MessageBox.Show("Failed to clear Exchange Calendar. Try Again .");
                }
            }
            searchSubscribersAndSelectRow(txtboxSearchFirstname.Text, txtboxSearchLastname.Text, comboBoxSex.Text, comboBoxTitle.Text, comboBoxOccupation.Text, txtboxSearchMainIdentifier.Text);
        }

        private void btnDeleteExchangeCalendar_Click(object sender, EventArgs e)
        {
            bool isSuccess = false;

            if (c.ExchangeCalendarID != null)
            {


                isSuccess = c.DeleteExchangeCalendar(c);
                if (isSuccess == true)
                {
                    MessageBox.Show("The selected Exchange Calendar has been successfuly deleted.");
                }
                else
                {
                    MessageBox.Show("Failed to delete Exchange Calendar. Try Again .");
                }
            }
            searchSubscribersAndSelectRow(txtboxSearchFirstname.Text, txtboxSearchLastname.Text, comboBoxSex.Text, comboBoxTitle.Text, comboBoxOccupation.Text, txtboxSearchMainIdentifier.Text);
        }
    }

    
    //
}
