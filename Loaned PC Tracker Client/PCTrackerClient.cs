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
        private bool Changed;
        private bool WindowLoaded;
        private TcpClient ClientSocket = new TcpClient() {
            NoDelay = true,
        };
        private NetworkStream ServerStream;
        private PCPacket ReceivePacket;

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

        private void UpdateStatus(string message) {
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
                ClientSocket.GetStream().Read(inStream, 0, ClientSocket.ReceiveBufferSize);
                List<string> sites = DeserializeStringStream(inStream);
                foreach(string s in sites) {
                    siteList.Add(s);
                }
            } catch (Exception ex) {
                bgwLoadSites.ReportProgress(0, ex.Message);
                //Console.WriteLine(" >> " + ex.Message.ToString());
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
            UpdateStatus((string)e.UserState);
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
        }
        
        /// <summary>
        ///     
        /// </summary>
        /// <param name="siteName"></param>
        /// <returns></returns>
        private string SetPCFileName(string siteName) {
            string[] splitSelection = siteName.Split(' ');
            return splitSelection[0] + ".xlsx";
        }

        /// <summary>
        ///     
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cbSiteChooser_SelectedIndexChanged(object sender, EventArgs e) {
            //PCFileName = SetPCFileName((string)cbSiteChooser.SelectedItem);
            rbHidden.Checked = true;
            btnSetDefaultSite.Enabled = true;
            CurrentlyAvailable.Clear();
            CheckedOut.Clear();
        }

        /// <summary>
        ///     
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSetDefaultSite_Click(object sender, EventArgs e) {
            SetDefaultSite(ProgramKey);
            int index = 0;// cbSiteChooser.FindString(PCFileName.Split('.')[0]);
            cbSiteChooser.SelectedIndex = index;
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
                if (Changed) {
                    using (var form = new ConfirmChanges()) {
                        var result = form.ShowDialog();
                        if (result == DialogResult.OK) {
                            SaveChanges(((string)cbSiteChooser.SelectedItem).Split(' ')[0], true);
                        }
                    }
                    Changed = false;
                }
                CurrentlyAvailable.Clear();
                CheckedOut.Clear();
                AccessLoanedPCData(((string)cbSiteChooser.SelectedItem).Split(' ')[0], false);
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
                if (Changed) {
                    using (var form = new ConfirmChanges()) {
                        var result = form.ShowDialog();
                        if (result == DialogResult.OK) {
                            SaveChanges(((string)cbSiteChooser.SelectedItem).Split(' ')[0], false);
                        }
                    }
                    Changed = false;
                }
                CurrentlyAvailable.Clear();
                CheckedOut.Clear();
                AccessLoanedPCData(((string)cbSiteChooser.SelectedItem).Split(' ')[0], true);
            }
        }

        /// <summary>
        ///     
        /// </summary>
        /// <param name="siteName"></param>
        /// <param name="hotswaps"></param>
        private void AccessLoanedPCData(string siteName, bool hotswaps) {
            
            //bgwLoadPCs.RunWorkerAsync(localFile);
            //ProgressBarForm = new LoadingProgress("Loading " + type + " List");
            //ProgressBarForm.ShowDialog();
        }

        /// <summary>
        ///     
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bgwLoadPCs_DoWork(object sender, DoWorkEventArgs e) {
            string fileName = (string)e.Argument;
            Excel.Workbook workbook = excelApp.Workbooks.Open(fileName);
            Excel.Worksheet currentSheet = workbook.Worksheets.Item[1];

            int lastRow = getMaxRow(currentSheet);
            int lastCol = getMaxCol(currentSheet);
            ProgressMax = lastRow;

            Laptop newLaptop;
            Laptop prevLaptop = new Laptop();

            for (int index = 2; index <= lastRow; index++) {
                // this array holds all of the information from each line of the excel sheet
                Array laptopValues = (Array)currentSheet.get_Range("A" + index.ToString(), ColumnNumToString(lastCol) + index.ToString()).Cells.Value;
                // I have to run the check null on each of these parsed cells, 
                // due to being brought in from an excel sheet with possible blank cells
                newLaptop = new Laptop() {
                    Number = intCheckNull(laptopValues.GetValue(1, 1)),
                    Serial = stringCheckNull(laptopValues.GetValue(1, 2)),
                    Brand = stringCheckNull(laptopValues.GetValue(1, 3)),
                    Model = stringCheckNull(laptopValues.GetValue(1, 4)),
                    Warranty = stringCheckNull(laptopValues.GetValue(1, 5)),
                    Username = stringCheckNull(laptopValues.GetValue(1, 6)),

                    TicketNumber = stringCheckNull(laptopValues.GetValue(1, 8)),
                    CheckedOut = boolCheckNull(laptopValues.GetValue(1, 9))
                };
                //this verifies that the newly created laptop is not a copy of the previous one
                if (newLaptop != prevLaptop) {
                    bgwLoadPCs.ReportProgress(index, newLaptop);
                    prevLaptop = newLaptop;
                }
            }
            workbook.Close();
        }

        /// <summary>
        ///     
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bgwLoadPCs_ProgressChanged(object sender, ProgressChangedEventArgs e) {
            Laptop sent = e.UserState as Laptop;
            if (sent.CheckedOut) {
                CheckedOut.Add(sent);
            } else {
                CurrentlyAvailable.Add(sent);
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
                        Changed = true;
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
                        Changed = true;
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
                    Changed = true;
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
                    Changed = true;
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
                    Changed = true;
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
            Changed = false;

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
    }
}
