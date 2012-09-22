using System;
using System.Globalization;
using System.Windows.Forms;
using Microsoft.Win32;
using System.IO;
using TimeSet;

namespace Schedule
{

    public class ScheduleForm : Form
    {
        Label lblView;
        ListView listView;
        ColumnHeader ScheduleName;
        ColumnHeader ScheduleFile;
        ColumnHeader ScheduleTime;
        Label lblName;
        TextBox txtName;
        Label lblFile;
        TextBox txtFile;
        Label lblTime;
        Button btnBrowse;
        OpenFileDialog openFileDialog;
        TimeSetControl timeSetControl;

        /// <summary>
        /// Required designer variable.
        /// </summary>
        System.ComponentModel.Container components = null;

        /// <summary>
        /// column headers for recreating the listbox
        /// </summary>
        ColumnHeader columnOne;
        ColumnHeader columnTwo;
        Button btnRemove;
        Button btnEdit;
        Button btnAdd;
        ColumnHeader columnThree;

        /// <summary>
        /// get and set the executable name
        /// </summary>
        public string Executable { get; set; }

        /// <summary>
        /// Construct the dialog and setup the list view columns
        /// </summary>
        public ScheduleForm()
        {
            //
            // Required for Windows Form Designer support
            //
            InitializeComponent();

            //
            // TODO: Add any constructor code after InitializeComponent call
            //
            UpdateColumns();

            UpdateListBox();

        }

        void UpdateColumns()
        {
            columnOne = new ColumnHeader();
            columnOne.Text = "Schedule Name";
            columnOne.Width = 150;

            columnTwo = new ColumnHeader();
            columnTwo.Text = "File";
            columnTwo.Width = 150;

            columnThree = new ColumnHeader();
            columnThree.Text = "Time";
            columnThree.Width = 150;
        }

        /// <summary>
        /// Function to ensure that the list view is kept up to date 
        /// called on initialisation and after add and remove functions
        /// </summary>
        public void UpdateListBox()
        {
            // clear the list box first
            listView.Clear();
            listView.Columns.Add(columnOne);
            listView.Columns.Add(columnTwo);
            listView.Columns.Add(columnThree);

            var regSoft = Registry.LocalMachine.OpenSubKey("Software", false);
            if (null == regSoft) return;

            var regSchedule = regSoft.OpenSubKey("ScheduleExample", false);
            if (null == regSchedule)
            {
                MessageBox.Show("Error unable to create the registry key 'ScheduleExample'\nThe service needs to be installed before this app will run");
                return;
            }

            var subKeys = regSchedule.GetSubKeyNames();
            var dtNow = DateTime.Now;

            foreach (var key in subKeys)
            {
                var regItem = regSchedule.OpenSubKey(key);
                if (null == regItem) continue;

                // get the value from the subkey
                var exeName = regItem.GetValue("FileToRun").ToString();
                var hours = Convert.ToInt32(regItem.GetValue("Hours").ToString());
                var mins = Convert.ToInt32(regItem.GetValue("Mins").ToString());

                var dtSchedule = new DateTime(dtNow.Year, dtNow.Month, dtNow.Day, hours, mins, 0, 0);

                var lvItem = listView.Items.Add(key);
                lvItem.SubItems.Add(exeName);
                lvItem.SubItems.Add(dtSchedule.ToString(CultureInfo.InvariantCulture));
            }
        }

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (components != null)
                {
                    components.Dispose();
                }
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code
        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        void InitializeComponent()
        {
            this.lblView = new Label();
            this.listView = new ListView();
            this.ScheduleName = ((ColumnHeader) (new ColumnHeader()));
            this.ScheduleFile = ((ColumnHeader) (new ColumnHeader()));
            this.ScheduleTime = ((ColumnHeader) (new ColumnHeader()));
            this.lblName = new Label();
            this.txtName = new TextBox();
            this.lblFile = new Label();
            this.txtFile = new TextBox();
            this.lblTime = new Label();
            this.timeSetControl = new TimeSet.TimeSetControl();
            this.btnBrowse = new Button();
            this.btnAdd = new Button();
            this.openFileDialog = new OpenFileDialog();
            this.btnRemove = new Button();
            this.btnEdit = new Button();
            this.SuspendLayout();
            // 
            // lblView
            // 
            this.lblView.Location = new System.Drawing.Point(8, 160);
            this.lblView.Name = "lblView";
            this.lblView.Size = new System.Drawing.Size(192, 23);
            this.lblView.TabIndex = 1;
            this.lblView.Text = "Currently Scheduled Items";
            // 
            // listView
            // 
            this.listView.Anchor = ((AnchorStyles) ((((AnchorStyles.Top | AnchorStyles.Bottom)
            | AnchorStyles.Left)
            | AnchorStyles.Right)));
            this.listView.Columns.AddRange(new ColumnHeader[] {
            this.ScheduleName,
            this.ScheduleFile,
            this.ScheduleTime});
            this.listView.FullRowSelect = true;
            this.listView.Location = new System.Drawing.Point(8, 184);
            this.listView.MultiSelect = false;
            this.listView.Name = "listView";
            this.listView.Size = new System.Drawing.Size(440, 104);
            this.listView.TabIndex = 2;
            this.listView.UseCompatibleStateImageBehavior = false;
            this.listView.View = View.Details;
            // 
            // ScheduleName
            // 
            this.ScheduleName.Text = "Schedule Name";
            this.ScheduleName.Width = 100;
            // 
            // ScheduleFile
            // 
            this.ScheduleFile.Text = "File";
            this.ScheduleFile.Width = 200;
            // 
            // ScheduleTime
            // 
            this.ScheduleTime.Text = "Time";
            this.ScheduleTime.Width = 100;
            // 
            // lblName
            // 
            this.lblName.Location = new System.Drawing.Point(8, 8);
            this.lblName.Name = "lblName";
            this.lblName.Size = new System.Drawing.Size(100, 16);
            this.lblName.TabIndex = 3;
            this.lblName.Text = "Schedule Name";
            // 
            // txtName
            // 
            this.txtName.Anchor = ((AnchorStyles) ((((AnchorStyles.Top | AnchorStyles.Bottom)
            | AnchorStyles.Left)
            | AnchorStyles.Right)));
            this.txtName.Location = new System.Drawing.Point(8, 24);
            this.txtName.Name = "txtName";
            this.txtName.Size = new System.Drawing.Size(120, 20);
            this.txtName.TabIndex = 4;
            // 
            // lblFile
            // 
            this.lblFile.Location = new System.Drawing.Point(8, 56);
            this.lblFile.Name = "lblFile";
            this.lblFile.Size = new System.Drawing.Size(128, 16);
            this.lblFile.TabIndex = 5;
            this.lblFile.Text = "File To Schedule";
            // 
            // txtFile
            // 
            this.txtFile.Anchor = ((AnchorStyles) ((((AnchorStyles.Top | AnchorStyles.Bottom)
            | AnchorStyles.Left)
            | AnchorStyles.Right)));
            this.txtFile.Location = new System.Drawing.Point(8, 72);
            this.txtFile.Name = "txtFile";
            this.txtFile.Size = new System.Drawing.Size(160, 20);
            this.txtFile.TabIndex = 6;
            // 
            // lblTime
            // 
            this.lblTime.Location = new System.Drawing.Point(8, 112);
            this.lblTime.Name = "lblTime";
            this.lblTime.Size = new System.Drawing.Size(80, 16);
            this.lblTime.TabIndex = 7;
            this.lblTime.Text = "Time To Run";
            // 
            // timeSetControl
            // 
            this.timeSetControl.Hours = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.timeSetControl.Location = new System.Drawing.Point(8, 128);
            this.timeSetControl.Minutes = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.timeSetControl.Name = "timeSetControl";
            this.timeSetControl.Seconds = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.timeSetControl.ShowHours = true;
            this.timeSetControl.ShowHoursLetter = true;
            this.timeSetControl.ShowMinutes = true;
            this.timeSetControl.ShowMinutesLetter = true;
            this.timeSetControl.ShowSeconds = false;
            this.timeSetControl.ShowSecondsLetter = false;
            this.timeSetControl.Size = new System.Drawing.Size(160, 32);
            this.timeSetControl.TabIndex = 8;
            // 
            // btnBrowse
            // 
            this.btnBrowse.Anchor = ((AnchorStyles) ((AnchorStyles.Top | AnchorStyles.Right)));
            this.btnBrowse.Location = new System.Drawing.Point(200, 72);
            this.btnBrowse.Name = "btnBrowse";
            this.btnBrowse.Size = new System.Drawing.Size(75, 23);
            this.btnBrowse.TabIndex = 11;
            this.btnBrowse.Text = "Browse";
            this.btnBrowse.Click += new System.EventHandler(this.OnBrowse);
            // 
            // btnAdd
            // 
            this.btnAdd.Anchor = ((AnchorStyles) ((AnchorStyles.Top | AnchorStyles.Right)));
            this.btnAdd.Location = new System.Drawing.Point(368, 8);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(75, 23);
            this.btnAdd.TabIndex = 10;
            this.btnAdd.Text = "Add";
            this.btnAdd.Click += new System.EventHandler(this.OnAdd);
            // 
            // btnRemove
            // 
            this.btnRemove.Anchor = ((AnchorStyles) ((AnchorStyles.Top | AnchorStyles.Right)));
            this.btnRemove.Location = new System.Drawing.Point(368, 40);
            this.btnRemove.Name = "btnRemove";
            this.btnRemove.Size = new System.Drawing.Size(75, 23);
            this.btnRemove.TabIndex = 12;
            this.btnRemove.Text = "Remove";
            this.btnRemove.Click += new System.EventHandler(this.OnRemoveButton);
            // 
            // btnEdit
            // 
            this.btnEdit.Anchor = ((AnchorStyles) ((AnchorStyles.Top | AnchorStyles.Right)));
            this.btnEdit.Location = new System.Drawing.Point(368, 72);
            this.btnEdit.Name = "btnEdit";
            this.btnEdit.Size = new System.Drawing.Size(75, 23);
            this.btnEdit.TabIndex = 13;
            this.btnEdit.Text = "Edit";
            this.btnEdit.Click += new System.EventHandler(this.OnEdit);
            // 
            // ScheduleForm
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(456, 302);
            this.Controls.Add(this.btnEdit);
            this.Controls.Add(this.btnRemove);
            this.Controls.Add(this.btnAdd);
            this.Controls.Add(this.btnBrowse);
            this.Controls.Add(this.timeSetControl);
            this.Controls.Add(this.lblTime);
            this.Controls.Add(this.txtFile);
            this.Controls.Add(this.lblFile);
            this.Controls.Add(this.txtName);
            this.Controls.Add(this.lblName);
            this.Controls.Add(this.listView);
            this.Controls.Add(this.lblView);
            this.Name = "ScheduleForm";
            this.Text = "Schedule Example";
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        #endregion

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.Run(new ScheduleForm());
        }

        /// <summary>
        /// Select a file to be used as the executable
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void OnBrowse(Object sender, EventArgs e)
        {
            openFileDialog.InitialDirectory = Directory.GetCurrentDirectory();
            openFileDialog.Filter = "Executable Files ( *.exe )| *.exe";
            openFileDialog.FilterIndex = 1;

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                Executable = openFileDialog.FileName;
                txtFile.Text = Executable;
            }
        }

        /// <summary>
        /// Add a new file to be scehduled
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void OnAdd(Object sender, EventArgs e)
        {
            btnAdd.Text = "Add";

            var reg = Registry.LocalMachine.OpenSubKey("Software", true).OpenSubKey("ScheduleExample", true);
            if (reg == null)
            {
                MessageBox.Show("Error unable to create the registry key 'ScheduleExample'");
                return;
            }

            var newKey = reg.CreateSubKey(txtName.Text);
            if (newKey == null)
            {
                MessageBox.Show("Error unable to create the registry key 'ScheduleExample'");
                return;
            }

            // put in some saftety code
            if (Executable == null)
            {
                MessageBox.Show("You must enter an executable to register");
                return;
            }

            newKey.SetValue("FileToRun", Executable);

            // note that the time control doesnt allow incorrect data
            // so there is no need to test it here.
            newKey.SetValue("Hours", timeSetControl.Hours.ToString());
            newKey.SetValue("Mins", timeSetControl.Minutes.ToString());

            UpdateListBox();

        }

        /// <summary>
        /// Remove the selected 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void OnRemoveButton(Object sender, EventArgs e)
        {
            var lvItemCol = listView.SelectedItems;
            if (0 == lvItemCol.Count)
            {
                MessageBox.Show("You must Select an item to remove");
                return;
            }

            // get the enumerator
            var colEnum = lvItemCol.GetEnumerator();

            // move to the first
            colEnum.MoveNext();

            // get the list view item
            var lvItem = (ListViewItem) colEnum.Current;

            // now remove the registry key
            var regSoft = Registry.LocalMachine.OpenSubKey("Software");
            if (null != regSoft)
            {
                var regSchedule = regSoft.OpenSubKey("ScheduleExample", true);
                if (null == regSchedule)
                {
                    MessageBox.Show("Unable to open the registry key");
                    return;
                }

                regSchedule.DeleteSubKeyTree(lvItem.Text);
            }

            UpdateListBox();
        }

        void OnEdit(Object sender, EventArgs e)
        {
            var lvItemCol = listView.SelectedItems;
            if (0 == lvItemCol.Count)
            {
                MessageBox.Show("You must Select an item to remove");
                return;
            }

            // get the enumerator
            var colEnum = lvItemCol.GetEnumerator();

            // move to the first
            colEnum.MoveNext();

            // get the list view item
            var lvItem = (ListViewItem) colEnum.Current;
            if (null == lvItem) throw new ArgumentNullException();

            var regSoft = Registry.LocalMachine.OpenSubKey("Software", true);
            if (null != regSoft)
            {
                var regSchedule = regSoft.OpenSubKey("ScheduleExample", true);
                if (null == regSchedule)
                {
                    MessageBox.Show("Error unable to create the registry key 'Sch	eduleExample'");
                    return;
                }

                var regItem = regSchedule.OpenSubKey(lvItem.Text);
                if (null == regItem)
                {
                    MessageBox.Show("Unable to open the key " + lvItem.Text);
                    return;
                }

                Executable = regItem.GetValue("FileToRun").ToString();
                timeSetControl.Hours = Convert.ToInt32(regItem.GetValue("Hours").ToString());
                timeSetControl.Minutes = Convert.ToInt32(regItem.GetValue("Mins").ToString());
            }

            txtName.Text = lvItem.Text;
            txtFile.Text = Executable;

            btnAdd.Text = "Update";
        }

    }
}
