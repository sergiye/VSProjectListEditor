using System;
using System.Windows.Forms;
using Microsoft.Win32;

namespace VSProjectListEditor {
    public class MainForm : Form {
        private Button btnSave;
        private Button btnUp;
        private Button btnDown;
        private Button btnDelete;
        private Button btnAdd;
        private OpenFileDialog openFileDialog;
        private ListBox lstBoxMain;

        private ComboBox cbDevStudioVersion;
        private RegistryKey mProjectMruKey;
        private Button btnRemoveOrphans;
        private Button btnRemoveNonSolutions;
        private Button btnExplorer;
        private Label lblVersion;

        private bool mBIsDirty;

        public MainForm() {
            InitializeComponent();

            // Add tool tips
            ToolTip ttAddNew = new ToolTip();
            ToolTip ttMoveUp = new ToolTip();
            ToolTip ttRemove = new ToolTip();
            ToolTip ttMoveDown = new ToolTip();
            ToolTip ttKilZombies = new ToolTip();
            ToolTip ttKilZomProj = new ToolTip();
            ToolTip ttExplorer = new ToolTip();
            ToolTip ttSave = new ToolTip();

            ttAddNew.SetToolTip(btnAdd, "Add a new entry");
            ttMoveUp.SetToolTip(btnUp, "Move selected entry up the list");
            ttRemove.SetToolTip(btnDelete, "Delete selected entry from the list");
            ttMoveDown.SetToolTip(btnDown, "Move selected entry down the list");
            ttKilZombies.SetToolTip(btnRemoveOrphans,
                    "Removes entries from the list which point to files that can't be found");
            ttKilZomProj.SetToolTip(btnRemoveNonSolutions,
                    "Remove all entries from the list except for Solution files");
            ttExplorer.SetToolTip(btnExplorer, "Open Explorer window in selected folder");
            ttSave.SetToolTip(btnSave, "Save the changes");

            // Add visual studio versions to list
            AddVersion("2010", @"Software\Microsoft\VisualStudio\10.0\ProjectMRUList");
            AddVersion("2008", @"Software\Microsoft\VisualStudio\9.0\ProjectMRUList");
            AddVersion("2005", @"Software\Microsoft\VisualStudio\8.0\ProjectMRUList");
            AddVersion("2005 Express", @"Software\Microsoft\VisualStudio\8.0Exp\ProjectMRUList");
            AddVersion("2003", @"Software\Microsoft\VisualStudio\7.1\ProjectMRUList");
            AddVersion("2002", @"Software\Microsoft\VisualStudio\7.0\ProjectMRUList");

            // select them until we find a non-empty list
            for (int i = 0; i < cbDevStudioVersion.Items.Count; i++) {
                cbDevStudioVersion.SelectedIndex = i;
                if (lstBoxMain.Items.Count != 0)
                    break;
            }
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        protected override void Dispose(bool disposing) {
            if (disposing) {
            }
            base.Dispose(disposing);
        }

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent() {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.btnSave = new System.Windows.Forms.Button();
            this.lstBoxMain = new System.Windows.Forms.ListBox();
            this.btnRemoveOrphans = new System.Windows.Forms.Button();
            this.cbDevStudioVersion = new System.Windows.Forms.ComboBox();
            this.btnDelete = new System.Windows.Forms.Button();
            this.btnDown = new System.Windows.Forms.Button();
            this.btnUp = new System.Windows.Forms.Button();
            this.btnAdd = new System.Windows.Forms.Button();
            this.btnRemoveNonSolutions = new System.Windows.Forms.Button();
            this.btnExplorer = new System.Windows.Forms.Button();
            this.lblVersion = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // openFileDialog
            // 
            this.openFileDialog.Filter = "All VS Files (.sln; .csproj)|*.sln;*.csproj;*.vbproj|All Project Files (*.csproj;" +
                                         "*.vbproj)|*.csproj;*.vbproj|Solution Files (*.sln)|*.sln|All Files (*.*)|*.*";
            // 
            // btnSave
            // 
            this.btnSave.Anchor =
                    ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSave.Location = new System.Drawing.Point(344, 338);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 23);
            this.btnSave.TabIndex = 9;
            this.btnSave.Text = "&Save";
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // lstBoxMain
            // 
            this.lstBoxMain.Anchor =
                    ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) |
                                                           System.Windows.Forms.AnchorStyles.Left) |
                                                          System.Windows.Forms.AnchorStyles.Right)));
            this.lstBoxMain.HorizontalScrollbar = true;
            this.lstBoxMain.Location = new System.Drawing.Point(12, 51);
            this.lstBoxMain.Name = "lstBoxMain";
            this.lstBoxMain.Size = new System.Drawing.Size(407, 277);
            this.lstBoxMain.TabIndex = 5;
            this.lstBoxMain.SelectedIndexChanged += new System.EventHandler(this.lstBoxMain_SelectedIndexChanged);
            this.lstBoxMain.DoubleClick += new System.EventHandler(this.lstBoxMain_DoubleClick);
            this.lstBoxMain.KeyDown += new System.Windows.Forms.KeyEventHandler(this.lstBoxMain_KeyDown);
            // 
            // btnRemoveOrphans
            // 
            this.btnRemoveOrphans.Anchor =
                    ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnRemoveOrphans.Location = new System.Drawing.Point(12, 338);
            this.btnRemoveOrphans.Name = "btnRemoveOrphans";
            this.btnRemoveOrphans.Size = new System.Drawing.Size(106, 23);
            this.btnRemoveOrphans.TabIndex = 6;
            this.btnRemoveOrphans.Text = "Remove &Orphans";
            this.btnRemoveOrphans.UseVisualStyleBackColor = true;
            this.btnRemoveOrphans.Click += new System.EventHandler(this.btnRemoveOrphans_Click);
            // 
            // cbDevStudioVersion
            // 
            this.cbDevStudioVersion.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbDevStudioVersion.FormattingEnabled = true;
            this.cbDevStudioVersion.Location = new System.Drawing.Point(12, 24);
            this.cbDevStudioVersion.Name = "cbDevStudioVersion";
            this.cbDevStudioVersion.Size = new System.Drawing.Size(217, 21);
            this.cbDevStudioVersion.TabIndex = 0;
            this.cbDevStudioVersion.SelectedIndexChanged += new System.EventHandler(this.cbDevStudioVersion_SelectedIndexChanged);
            // 
            // btnDelete
            // 
            this.btnDelete.Anchor =
                    ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnDelete.Image = ((System.Drawing.Image)(resources.GetObject("btnDelete.Image")));
            this.btnDelete.Location = new System.Drawing.Point(395, 22);
            this.btnDelete.Name = "btnDelete";
            this.btnDelete.Size = new System.Drawing.Size(24, 23);
            this.btnDelete.TabIndex = 4;
            this.btnDelete.UseVisualStyleBackColor = true;
            this.btnDelete.Click += new System.EventHandler(this.btnDelete_Click);
            // 
            // btnDown
            // 
            this.btnDown.Anchor =
                    ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnDown.Image = ((System.Drawing.Image)(resources.GetObject("btnDown.Image")));
            this.btnDown.Location = new System.Drawing.Point(365, 22);
            this.btnDown.Name = "btnDown";
            this.btnDown.Size = new System.Drawing.Size(24, 23);
            this.btnDown.TabIndex = 3;
            this.btnDown.Click += new System.EventHandler(this.btnDown_Click);
            // 
            // btnUp
            // 
            this.btnUp.Anchor =
                    ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnUp.Image = ((System.Drawing.Image)(resources.GetObject("btnUp.Image")));
            this.btnUp.Location = new System.Drawing.Point(335, 22);
            this.btnUp.Name = "btnUp";
            this.btnUp.Size = new System.Drawing.Size(24, 23);
            this.btnUp.TabIndex = 2;
            this.btnUp.Click += new System.EventHandler(this.btnUp_Click);
            // 
            // btnAdd
            // 
            this.btnAdd.Anchor =
                    ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnAdd.Image = ((System.Drawing.Image)(resources.GetObject("btnAdd.Image")));
            this.btnAdd.Location = new System.Drawing.Point(305, 22);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(24, 23);
            this.btnAdd.TabIndex = 1;
            this.btnAdd.UseVisualStyleBackColor = true;
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // btnRemoveNonSolutions
            // 
            this.btnRemoveNonSolutions.Anchor =
                    ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnRemoveNonSolutions.Location = new System.Drawing.Point(124, 338);
            this.btnRemoveNonSolutions.Name = "btnRemoveNonSolutions";
            this.btnRemoveNonSolutions.Size = new System.Drawing.Size(105, 23);
            this.btnRemoveNonSolutions.TabIndex = 7;
            this.btnRemoveNonSolutions.Text = "Sol&utions Only";
            this.btnRemoveNonSolutions.UseVisualStyleBackColor = true;
            this.btnRemoveNonSolutions.Click += new System.EventHandler(this.btnRemoveNonSolutions_Click);
            // 
            // btnExplorer
            // 
            this.btnExplorer.Anchor =
                    ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnExplorer.Location = new System.Drawing.Point(235, 338);
            this.btnExplorer.Name = "btnExplorer";
            this.btnExplorer.Size = new System.Drawing.Size(103, 23);
            this.btnExplorer.TabIndex = 8;
            this.btnExplorer.Text = "&Explorer...";
            this.btnExplorer.UseVisualStyleBackColor = true;
            this.btnExplorer.Click += new System.EventHandler(this.btnExplorer_Click);
            // 
            // lblVersion
            // 
            this.lblVersion.AutoSize = true;
            this.lblVersion.Location = new System.Drawing.Point(13, 8);
            this.lblVersion.Name = "lblVersion";
            this.lblVersion.Size = new System.Drawing.Size(105, 13);
            this.lblVersion.TabIndex = 10;
            this.lblVersion.Text = "Visual Studio &version";
            // 
            // MainForm
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(431, 372);
            this.Controls.Add(this.cbDevStudioVersion);
            this.Controls.Add(this.lblVersion);
            this.Controls.Add(this.btnExplorer);
            this.Controls.Add(this.btnRemoveNonSolutions);
            this.Controls.Add(this.btnRemoveOrphans);
            this.Controls.Add(this.btnDown);
            this.Controls.Add(this.btnDelete);
            this.Controls.Add(this.lstBoxMain);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.btnUp);
            this.Controls.Add(this.btnAdd);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MinimumSize = new System.Drawing.Size(439, 382);
            this.Name = "MainForm";
            this.Text = "VS Project List Editor";
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.Closed += new System.EventHandler(this.MainForm_Closed);
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.MainForm_FormClosing);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        #region Main

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main() {
            Application.Run(new MainForm());
        }

        #endregion

        #region Button Events

        /// <summary>
        /// Save (Writes) Registry Values from ListBox
        /// </summary>
        private void btnSave_Click(object sender, EventArgs e) {
            Save();
        }

        /// <summary>
        /// Deletes a value in the ListBox
        /// </summary>
        private void btnDelete_Click(object sender, EventArgs e) {
            DeleteSelected();
        }

        private void DeleteSelected() {
            int index = lstBoxMain.SelectedIndex;
            if (index < 0) return;
            lstBoxMain.Items.RemoveAt(index);
            if (lstBoxMain.Items.Count > 0)
                lstBoxMain.SelectedIndex = Math.Min(index, lstBoxMain.Items.Count - 1); // becomes -1 if lb is empty
            SetDirty(true);
        }

        /// <summary>
        /// Moves a value up in the listbox
        /// </summary>
        private void btnUp_Click(object sender, EventArgs e) {
            MoveUp();
        }

        private void MoveUp() {
            int index = lstBoxMain.SelectedIndex;
            if (index <= 0) return;
            lstBoxMain.Items.Insert(index - 1, lstBoxMain.SelectedItem);
            lstBoxMain.SelectedIndex = index - 1;
            lstBoxMain.Items.RemoveAt(index + 1);
            SetDirty(true);
        }

        /// <summary>
        /// Moves a value down in the ListBox
        /// </summary>
        private void btnDown_Click(object sender, EventArgs e) {
            MoveDown();
        }

        private void MoveDown() {
            int index = lstBoxMain.SelectedIndex;
            if (index >= lstBoxMain.Items.Count - 1 || index < 0) return;
            lstBoxMain.Items.Insert(index + 2, lstBoxMain.SelectedItem);
            lstBoxMain.SelectedIndex = index + 2;
            lstBoxMain.Items.RemoveAt(index);
            SetDirty(true);
        }

        /// <summary>
        /// Adds a new file path to the list box
        /// </summary>
        private void btnAdd_Click(object sender, EventArgs e) {
            AddProject();
        }

        private void AddProject() {
            if (openFileDialog.ShowDialog() != DialogResult.OK) return;
            lstBoxMain.Items.Insert(0, openFileDialog.FileName);
            lstBoxMain.SelectedIndex = 0;
            SetDirty(true);
        }

        /// <summary>
        /// Removes entries which point to files that can't be found
        /// </summary>
        private void btnRemoveOrphans_Click(object sender, EventArgs e) {
            bool bRemoved = false;

            for (int idx = lstBoxMain.Items.Count - 1; idx >= 0; --idx) {
                string path = lstBoxMain.Items[idx].ToString();
                if (System.IO.File.Exists(path)) continue;
                lstBoxMain.Items.RemoveAt(idx);
                bRemoved = true;
            }

            // Don't use SetDirty(bRemoved) as we don't want to set
            // the flag to 'clean' if no entries removed!
            if (bRemoved)
                SetDirty(true);
        }

        /// <summary>
        /// Removes entries for files other than Solutions
        /// </summary>
        private void btnRemoveNonSolutions_Click(object sender, EventArgs e) {
            bool bRemoved = false;

            for (int idx = lstBoxMain.Items.Count - 1; idx >= 0; --idx) {
                string path = lstBoxMain.Items[idx].ToString();
                // Path.GetExtension complains about invalid characters, which the
                // registry key sometimes contains, so manually:
                string ext = "";
                int ppos = path.LastIndexOf('.');
                if (ppos >= 0)
                    ext = path.Substring(ppos + 1);
                if (ext.ToLower() == "sln") continue;
                lstBoxMain.Items.RemoveAt(idx);
                bRemoved = true;
            }

            // Don't use SetDirty(bRemoved) as we don't want to set
            // the flag to 'clean' if no entries removed!
            if (bRemoved)
                SetDirty(true);
        }

        /// <summary>
        /// Open an Explorer window in the folder of the selected file
        /// </summary>
        private void btnExplorer_Click(object sender, EventArgs e) {
            BrowseToFile();
        }

        #endregion

        #region Form Events

        private void MainForm_Load(object sender, EventArgs e) {
            if (cbDevStudioVersion.Items.Count == 0) {
                MessageBox.Show("No Visual Studio MRU lists were found!",
                        "Error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Exclamation);
            }
        }

        private void MainForm_Closed(object sender, EventArgs e) {
            //Close and flush the registry value on close
            if (mProjectMruKey != null)
                mProjectMruKey.Close();
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e) {
            if (mProjectMruKey != null && mBIsDirty && AskSaveChanges()) {
                Save();
            }
        }

        #endregion

        #region Other control events

        private void cbDevStudioVersion_SelectedIndexChanged(object sender, EventArgs e) {
            TaggedString<string> sel = cbDevStudioVersion.SelectedItem as TaggedString<string>;

            if (mProjectMruKey != null) {
                if (mBIsDirty && AskSaveChanges())
                    Save();
                else
                    SetDirty(false);

                mProjectMruKey.Close();
                mProjectMruKey = null;
            }

            if (sel != null) {
                mProjectMruKey = Registry.CurrentUser.OpenSubKey(sel.Tag, true);
            }
            else {
                if (mProjectMruKey != null)
                    mProjectMruKey.Close();
                mProjectMruKey = null;
            }

            UpdateUI();
        }

        /// <summary>
        ///  Enable or disable buttons depending on selection
        /// </summary>
        private void lstBoxMain_SelectedIndexChanged(object sender, EventArgs e) {
            int iSelected = ((ListBox)sender).SelectedIndex;

            btnUp.Enabled = (iSelected != 0);
            btnDown.Enabled = (iSelected != ((ListBox)sender).Items.Count - 1);
            btnDelete.Enabled = true;
            btnExplorer.Enabled = true;
        }

        private void lstBoxMain_DoubleClick(object sender, EventArgs e) {
            BrowseToFile();
        }

        #endregion

        #region Private events

        private void Save() {
            if (mProjectMruKey == null)
                return;

            foreach (string key in mProjectMruKey.GetValueNames())
                mProjectMruKey.DeleteValue(key);

            for (int i = 0; i < lstBoxMain.Items.Count; ++i) {
                mProjectMruKey.SetValue("File" + (i + 1), lstBoxMain.Items[i]);
            }
            //Flush the object to force it to write the files to the Registry
            // m_ProjectMRUKey.Flush(); - Usually not necessary (physical flush of the registry to disk)
            SetDirty(false);
        }

        private void UpdateUI() {
            lstBoxMain.Items.Clear();
            if (mProjectMruKey != null) {
                foreach (string key in mProjectMruKey.GetValueNames())
                    lstBoxMain.Items.Add(mProjectMruKey.GetValue(key));
            }

            lstBoxMain.Enabled = mProjectMruKey != null;

            btnUp.Enabled = false;
            btnDown.Enabled = false;
            btnDelete.Enabled = false;
            btnExplorer.Enabled = false;
            btnSave.Enabled = false;
        }

        /// <summary>
        /// Open an Explorer window in the folder of the selected file
        /// </summary>
        private void BrowseToFile() {
            if (lstBoxMain.SelectedItem == null)
                return;
            string path = lstBoxMain.SelectedItem.ToString();

            System.Diagnostics.Process.Start("explorer.exe", "/e,/select,\"" + path + "\"");
        }

        /// <summary>
        /// Add a new entry to the versions list if it's installed
        /// </summary>
        /// <param name="name">Name as it will appear in the list</param>
        /// <param name="regKey">MRU registry key for this version</param>
        private void AddVersion(string name, string regKey) {
            // Only add it if we can find the key
            if (Registry.CurrentUser.OpenSubKey(regKey) != null)
                cbDevStudioVersion.Items.Add(new TaggedString<string>(name, regKey));
        }

        /// <summary>
        /// Set the 'dirty' flag and enable or disable save button
        /// </summary>
        /// <param name="isDirty">True if changes have been made, False otherwise</param>
        private void SetDirty(bool isDirty) {
            mBIsDirty = isDirty;
            btnSave.Enabled = isDirty;
        }

        /// <summary>
        /// Ask whether user wants to save changes
        /// </summary>
        /// <returns>True if Yes, False if No</returns>
        private bool AskSaveChanges() {
            DialogResult drAnswer = MessageBox.Show(this,
                    "Save Changes?",
                    "Visual Studio Project Editor",
                    MessageBoxButtons.YesNo);

            return (drAnswer == DialogResult.Yes);
        }

        #endregion

        private void lstBoxMain_KeyDown(object sender, KeyEventArgs e) {
            switch (e.KeyData) {
                case Keys.Down | Keys.Shift:
                    MoveDown();
                    break;
                case Keys.Up | Keys.Shift:
                    MoveUp();
                    break;
                case Keys.Delete:
                    DeleteSelected();
                    break;
                case Keys.Add:
                    AddProject();
                    break;
                default: return; // skip "Handled = true"
            }
            e.Handled = true; // reset to false if no match was found
        }

    }
}