using System;
using System.ComponentModel;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Collections.Generic;
using System.Text;

namespace Loaned_PC_Tracker_Client {
    public partial class PCTrackerClient : Form {
        private const string KeyLocation = "SOFTWARE\\PC Tracker";
        private Microsoft.Win32.RegistryKey ProgramKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(KeyLocation);
        private string Server = "MXL3090GHT-X7";
        private BindingList<Laptop> CurrentlyAvailable = new BindingList<Laptop>();
        private BindingList<Laptop> CheckedOut = new BindingList<Laptop>();
        private BindingList<string> siteList = new BindingList<string>();
        private Excel.Application excelApp = new Excel.Application() {
            Visible = false,
            DisplayAlerts = false
        };
        private LoadingProgress ProgressBarForm;
        private int ProgressMax;
        private bool WindowLoaded;
        private TcpClient ClientSocket = new TcpClient() {
            NoDelay = true,
        };
        private delegate void StringParameterDelegate(string value);

        public PCTrackerClient() {
            InitializeComponent();

            cbSiteChooser.DataSource = siteList;
            dgvAvailable.DataSource = CurrentlyAvailable;
            dgvCheckedOut.DataSource = CheckedOut;
        }

        /// <summary>
        ///     
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void frmPCTracker_Activated(object sender, EventArgs e) {
            if (!WindowLoaded) {
                WindowLoaded = true;
                int numRetries = 1;
                using (var connectTo = new ConnectionForm(Server)) {
                    var result = connectTo.ShowDialog();
                    if (result == DialogResult.OK) {
                        Server = connectTo.ReturnAddress;
                        numRetries = connectTo.ReturnRetries;
                    } else {
                        Close();
                    }
                }
                ConnectToServer(numRetries);
                bgwLoadSites.RunWorkerAsync();
                ProgressBarForm = new LoadingProgress("Receiving Sites List");
                ProgressBarForm.ShowDialog();
            }
        }

        private void ConnectToServer(int retries) {
            //this loop allows for multiple attempts to connect to the server before timing out
            for (int i = 0; i < retries; i++) {
                UpdateStatus(">>Attempting to connect...");
                try {
                    ClientSocket.Connect(Server, 8888);
                } catch (Exception ex) {
                    UpdateStatus(ex.Message);
                }

                //once connected, the client sends a packet to the server to agknowledge logging in.
                if (ClientSocket.Connected) {
                    NetworkStream loginStream = ClientSocket.GetStream();
                    NamePacket handshake = new NamePacket(Environment.UserName);
                    loginStream.Write(handshake.CreateDataStream(), 0, handshake.PacketLength);
                    loginStream.Flush();
                    UpdateStatus("Server Connected ... " + Server);
                    //bwBroadcastStream.RunWorkerAsync();
                    break;
                }
                UpdateStatus(">> Waiting 3 Seconds before trying again.");
                DateTime start = DateTime.Now;
                while (DateTime.Now.Subtract(start).Seconds < 3) { }
            }
            if (!ClientSocket.Connected) {
                UpdateStatus("Connection Failed!!");
            }
        }

        public void UpdateStatus(string message) {
            if (InvokeRequired) {
                // We're not in the UI thread, so we need to call BeginInvoke
                BeginInvoke(new StringParameterDelegate(UpdateStatus), new object[] { message });
                return;
            }
            // Must be on the UI thread if we've got this far
            tbConnectionStatus.AppendText(message);
            tbConnectionStatus.AppendText(Environment.NewLine);
        }

        /// <summary>
        ///     
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void frmPCTracker_Resize(object sender, EventArgs e) {
            Form sent = sender as Form;
            if (sent.WindowState == FormWindowState.Maximized) {
                dgvAvailable.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                dgvCheckedOut.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            } else {
                dgvAvailable.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;
                dgvCheckedOut.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;
            }
        }
        
        /// <summary>
        ///     
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bgwLoadSites_DoWork(object sender, DoWorkEventArgs e) {
            
            byte[] inStream = new byte[10025];
            try {
                //ClientSocket.GetStream().Read(inStream, 0, ClientSocket.ReceiveBufferSize);
                //ProgressMax = DeserializeIntStream(inStream);
                ClientSocket.GetStream().Read(inStream, 0, ClientSocket.ReceiveBufferSize);
                List<string> sites = DeserializeStringStream(inStream);
                foreach(string s in sites) {
                    siteList.Add(s);
                }
            } catch (Exception ex) {
                UpdateStatus(ex.Message);
                //bgwLoadSites.ReportProgress(0, ex.Message);
            }
        }

        private List<string> DeserializeStringStream(byte[] stream) {
            string stringStream = Encoding.UTF8.GetString(stream);
            string[] splitStream = stringStream.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            return splitStream.ToList();
        }
        
        /// <summary>
        ///     
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bgwLoadSites_ProgressChanged(object sender, ProgressChangedEventArgs e) {
            //UpdateStatus((string)e.UserState);
            if (ProgressBarForm.getProgressMaximum() != ProgressMax) {
                ProgressBarForm.setProgressMaximum(ProgressMax);
            }
            ProgressBarForm.updateProgress(e.ProgressPercentage);
        }

        /// <summary>
        ///     
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bgwLoadSites_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e) {
            siteList.ResetBindings();

            int index = cbSiteChooser.FindString(GetDefaultSite(ProgramKey));
            cbSiteChooser.SelectedIndex = index;
            btnSetDefaultSite.Enabled = false;

            ProgressBarForm.Close();
            bgwAwaitBroadcasts.RunWorkerAsync();
        }

        /// <summary>
        ///     
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cbSiteChooser_SelectedIndexChanged(object sender, EventArgs e) {
            if (cbSiteChooser.SelectedIndex != cbSiteChooser.FindString(GetDefaultSite(ProgramKey))) {
                rbHidden.Checked = true;
                btnSetDefaultSite.Enabled = true;
                CurrentlyAvailable.Clear();
                CheckedOut.Clear();
            } else {
                btnSetDefaultSite.Enabled = false;
            }
        }

        /// <summary>
        ///     
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSetDefaultSite_Click(object sender, EventArgs e) {
            SetDefaultSite(ProgramKey);
            btnSetDefaultSite.Enabled = false;
        }

        /// <summary>
        ///     
        /// </summary>
        /// <param name="key"></param>
        /// <returns></returns>
        private string GetDefaultSite(Microsoft.Win32.RegistryKey key) {
            if (key == null) {
                SetDefaultSite(key);
                return (string)cbSiteChooser.SelectedItem;
            }
            return (string)key.GetValue("Default Site");
        }

        /// <summary>
        ///     
        /// </summary>
        /// <param name="key"></param>
        private void SetDefaultSite(Microsoft.Win32.RegistryKey key) {
            string siteName = (string)cbSiteChooser.SelectedItem;

            if (key == null) {
                key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(KeyLocation);
                key.SetValue("Default Site", siteName);
            } else {
                Microsoft.Win32.Registry.CurrentUser.DeleteSubKey(KeyLocation);
                key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(KeyLocation);
                key.SetValue("Default Site", siteName);
            }
            key.Close();
        }

        /// <summary>
        ///     
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void rbLoaner_CheckedChanged(object sender, EventArgs e) {
            RadioButton sent = sender as RadioButton;
            if (sent.Checked) {
                /*if (Changed) {
                    using (var form = new ConfirmChanges()) {
                        var result = form.ShowDialog();
                        if (result == DialogResult.OK) {
                            SaveChanges(((string)cbSiteChooser.SelectedItem).Split(' ')[0], true);
                        }
                    }
                    Changed = false;
                }*/
                CurrentlyAvailable.Clear();
                CheckedOut.Clear();
                AccessLoanedPCData((string)cbSiteChooser.SelectedItem, false);
            }
        }

        /// <summary>
        ///     
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void rbHotSwap_CheckedChanged(object sender, EventArgs e) {
            RadioButton sent = sender as RadioButton;
            if (sent.Checked) {
                /*if (Changed) {
                    using (var form = new ConfirmChanges()) {
                        var result = form.ShowDialog();
                        if (result == DialogResult.OK) {
                            SaveChanges(((string)cbSiteChooser.SelectedItem).Split(' ')[0], false);
                        }
                    }
                    Changed = false;
                }*/
                CurrentlyAvailable.Clear();
                CheckedOut.Clear();
                AccessLoanedPCData((string)cbSiteChooser.SelectedItem, true);
            }
        }

        /// <summary>
        ///     
        /// </summary>
        /// <param name="siteName"></param>
        /// <param name="hotswaps"></param>
        private void AccessLoanedPCData(string siteName, bool hotswaps) {
            string type = string.Empty;
            if (hotswaps) {
                type = "Hotswaps";
            } else {
                type = "Loaners";
            }

            RequestPCPacket requestPCs = new RequestPCPacket(siteName, type);
            bgwLoadPCs.RunWorkerAsync(requestPCs);
            ProgressBarForm = new LoadingProgress("Loading " + type + " List");
            ProgressBarForm.ShowDialog();
        }

        /// <summary>
        ///     
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bgwLoadPCs_DoWork(object sender, DoWorkEventArgs e) {
            var requestPCs = e.Argument as RequestPCPacket;
            byte[] inStream = new byte[10025];
            
            try {
                NetworkStream stream = ClientSocket.GetStream();
                UpdateStatus("Requesting PC's for " + requestPCs.SiteName);
                stream.Write(requestPCs.CreateDataStream(), 0, requestPCs.PacketLength);
                stream.Flush();
                stream.Read(inStream, 0, ClientSocket.ReceiveBufferSize);
                //List<Laptop> receivedPCs = SplitPCStream(inStream);
                SplitPCStream(inStream);
            } catch (Exception ex) {
                UpdateStatus(ex.Message);
            }
            
        }

        //private List<Laptop> SplitPCStream(byte[] dataStream) {
        private void SplitPCStream(byte[] dataStream) {
            var seperator = new char[] { ';' };
            var stringStream = Encoding.UTF8.GetString(dataStream);
            var splitStream = stringStream.Split(seperator, StringSplitOptions.RemoveEmptyEntries);
            //var returnList = new List<Laptop>();
            foreach (string s in splitStream) {
                bgwLoadPCs.ReportProgress(0, new Laptop().DeserializeLaptop(s));
                //returnList.Add(new Laptop().DeserializeLaptop(s));
            }
            //return returnList;
        }

        /// <summary>
        ///     
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bgwLoadPCs_ProgressChanged(object sender, ProgressChangedEventArgs e) {
            Laptop PC = e.UserState as Laptop;
            if (PC.CheckedOut) {
                CheckedOut.Add(PC);
            } else {
                CurrentlyAvailable.Add(PC);
            }

            if (ProgressBarForm.getProgressMaximum() != ProgressMax) {
                ProgressBarForm.setProgressMaximum(ProgressMax);
            }

            ProgressBarForm.updateProgress(e.ProgressPercentage);
        }

        /// <summary>
        ///     
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bgwLoadPCs_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e) {
            CurrentlyAvailable.ResetBindings();
            CheckedOut.ResetBindings();
            ProgressBarForm.Close();
        }

        /// <summary>
        ///     check for a blank cell value and return a string, if that is expected
        /// </summary>
        /// <param name="value"></param>
        /// <returns> a string of the cell contents</returns>
        private string stringCheckNull(object value) {
            if (value == null) {
                return "";
            }
            return value.ToString();
        }

        /// <summary>
        ///     check for a blank cell value and return a boolean, if that is expected
        /// </summary>
        /// <param name="value"></param>
        /// <returns> a if the cell contents are null or true/false</returns>
        private bool boolCheckNull(object value) {
            if (value == null) {
                return false;
            }
            if ((bool)value) {
                return true;
            }
            return false;
        }

        /// <summary>
        ///     check for a blank cell value and return an integer, if that is expected
        /// </summary>
        /// <param name="value"></param>
        /// <returns> 0 if the cell is blank or the object doesn't parse, otherwise returns an int value</returns>
        private int intCheckNull(object value) {
            int parsedNum;
            if (value == null) {
                return 0;
            }
            if (int.TryParse(value.ToString(), out parsedNum)) {
                return parsedNum;
            } else {
                return 0;
            }
        }

        /// <summary>
        ///     Returns the last row number that has any information in any cell of an excel sheet
        /// </summary>
        /// <param name="worksheet"></param>
        /// <returns> the last row number with any data </returns>
        private int getMaxRow(Excel.Worksheet worksheet) {
            int lastRow = worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            return lastRow;
        }

        /// <summary>
        ///     returns the last column number that has any information in any cell of an excel sheet
        /// </summary>
        /// <param name="worksheet"></param>
        /// <returns> the last column number with any data </returns>
        private int getMaxCol(Excel.Worksheet worksheet) {
            int lastCol = worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;
            return lastCol;
        }

        /// <summary>
        ///     returns a full column name using the column number as a basis
        /// </summary>
        /// <param name="columnNumber"></param>
        /// <returns></returns>
        private string ColumnNumToString(int columnNumber) {
            int dividend = columnNumber;
            string strColumnName = "";
            int modulo;
            while (dividend > 0) {
                modulo = (dividend - 1) % 26;
                strColumnName = Convert.ToChar(65 + modulo).ToString() + strColumnName;
                dividend = (int)((dividend - modulo) / 26);
            }
            return strColumnName;
        }

        /// <summary>
        ///     
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgvAvailable_CellClick(object sender, DataGridViewCellEventArgs e) {
            Viewer viewItem = new Viewer((Laptop)dgvAvailable.SelectedRows[0].DataBoundItem);
            viewItem.ShowDialog();
        }

        /// <summary>
        ///     
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgvCheckedOut_CellClick(object sender, DataGridViewCellEventArgs e) {
            Viewer viewItem = new Viewer((Laptop)dgvCheckedOut.SelectedRows[0].DataBoundItem);
            viewItem.ShowDialog();
        }

        /// <summary>
        ///     
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCheckOut_Click(object sender, EventArgs e) {
            Laptop checkOutPC = (Laptop)dgvAvailable.SelectedRows[0].DataBoundItem;
            if (checkOutPC != null) {
                using (var form = new CheckOutOrIn(checkOutPC, rbHotSwaps.Checked)) {
                    var result = form.ShowDialog();
                    if (result == DialogResult.OK) {
                        checkOutPC = form.ReturnPC;
                        CurrentlyAvailable.Remove(checkOutPC);
                        CheckedOut.Add(checkOutPC);
                        //Changed = true;
                    }
                }
            }
        }

        /// <summary>
        ///     
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCheckIn_Click(object sender, EventArgs e) {
            Laptop checkInPC = (Laptop)dgvCheckedOut.SelectedRows[0].DataBoundItem;
            if (checkInPC != null) {
                using (var form = new CheckOutOrIn(checkInPC, rbHotSwaps.Checked, true)) {
                    var result = form.ShowDialog();
                    if (result == DialogResult.OK) {
                        checkInPC = form.ReturnPC;
                        CheckedOut.Remove(checkInPC);
                        CurrentlyAvailable.Add(checkInPC);
                        //Changed = true;
                    }
                }
            }
        }

        /// <summary>
        ///     
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnEditPC_Click(object sender, EventArgs e) {
            Laptop editedPC = (Laptop)dgvAvailable.SelectedRows[0].DataBoundItem;
            using (var form = new AddEditRemove(editedPC, false, rbHotSwaps.Checked)) {
                var result = form.ShowDialog();
                if (result == DialogResult.OK) {
                    CurrentlyAvailable.Remove(editedPC);
                    editedPC = form.ReturnPC;
                    CurrentlyAvailable.Add(editedPC);
                    //Changed = true;
                }
            }
        }

        /// <summary>
        ///     
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnAddNew_Click(object sender, EventArgs e) {
            using (var form = new AddEditRemove(rbHotSwaps.Checked)) {
                var result = form.ShowDialog();
                if (result == DialogResult.OK) {
                    CurrentlyAvailable.Add(form.ReturnPC);
                    //Changed = true;
                }
            }
        }

        /// <summary>
        ///     
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnRemoveOld_Click(object sender, EventArgs e) {
            Laptop PCtoRemove = (Laptop)dgvAvailable.SelectedRows[0].DataBoundItem;
            using (var form = new AddEditRemove(PCtoRemove, true)) {
                var result = form.ShowDialog();
                if (result == DialogResult.OK) {
                    CurrentlyAvailable.Remove(PCtoRemove);
                    //Changed = true;
                }
            }
        }

        /// <summary>
        ///     
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSaveChanges_Click(object sender, EventArgs e) {
            bool hotswap = false;
            if (rbHotSwaps.Checked) {
                hotswap = true;
            }
            SaveChanges(((string)cbSiteChooser.SelectedItem).Split(' ')[0], hotswap);
        }

        /// <summary>
        ///     
        /// </summary>
        /// <param name="siteName"></param>
        /// <param name="hotswaps"></param>
        private void SaveChanges(string siteName, bool hotswaps) {
            //bgwSaveChanges.RunWorkerAsync(localFile);
            //ProgressBarForm = new LoadingProgress("Saving " + type + " List");
            //ProgressBarForm.ShowDialog();
        }

        /// <summary>
        ///     
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bgwSaveChanges_DoWork(object sender, DoWorkEventArgs e) {
            string fileName = (string)e.Argument;
            Excel.Workbook workbook = excelApp.Workbooks.Open(fileName);
            Excel.Worksheet currentSheet = workbook.Worksheets.Item[1];

            int lastrow = 2;
            int progress = 0;
            ProgressMax = (CurrentlyAvailable.Count + CheckedOut.Count);

            //todo: add logic to replace the data in the sheet with the new data
            foreach (Laptop PC in CurrentlyAvailable.Union(CheckedOut)) {
                currentSheet.Rows[lastrow].Delete();
                currentSheet.Cells[lastrow, 1].Value = PC.Number.ToString();
                currentSheet.Cells[lastrow, 2].Value = PC.Serial;
                currentSheet.Cells[lastrow, 3].Value = PC.Brand;
                currentSheet.Cells[lastrow, 4].Value = PC.Model;
                currentSheet.Cells[lastrow, 5].Value = PC.Warranty;
                currentSheet.Cells[lastrow, 6].Value = PC.Username;
                currentSheet.Cells[lastrow, 7].Value = PC.UserPCSerial;
                currentSheet.Cells[lastrow, 8].Value = PC.TicketNumber;
                currentSheet.Cells[lastrow, 9].Value = PC.CheckedOut;
                lastrow++;
                bgwSaveChanges.ReportProgress(++progress);
            }
            workbook.Save();
            workbook.Close();
        }

        /// <summary>
        ///     
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bgwSaveChanges_ProgressChanged(object sender, ProgressChangedEventArgs e) {
            if (ProgressBarForm.getProgressMaximum() != ProgressMax) {
                ProgressBarForm.setProgressMaximum(ProgressMax);
            }
            ProgressBarForm.updateProgress(e.ProgressPercentage);
        }

        /// <summary>
        ///     
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bgwSaveChanges_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e) {
            CurrentlyAvailable.ResetBindings();
            CheckedOut.ResetBindings();
            //Changed = false;

            ProgressBarForm.Close();
        }

        /// <summary>
        ///     
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void frmPCTracker_Closing(object sender, FormClosingEventArgs e) {
            excelApp.Quit();
        }

        private void bgwAwaitBroadcasts_DoWork(object sender, DoWorkEventArgs e) {
            NetworkStream broadcastStream;
            byte[] inStream = new byte[10025];
            while (true) {
                try {
                    broadcastStream = ClientSocket.GetStream();
                    broadcastStream.Read(inStream, 0, ClientSocket.ReceiveBufferSize);
                    List<string> broadcast = DeserializeStringStream(inStream);
                    foreach (string s in broadcast) {
                        UpdateStatus(s);
                    }
                } catch (Exception ex) {
                    UpdateStatus(ex.Message);
                    break;
                }
            }
        }

        private void bgwAwaitBroadcasts_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e) {
            UpdateStatus("Disconnected from server!");
        }
    }
}
