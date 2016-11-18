using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Net;
using System.Net.Sockets;

namespace Loaned_PC_Tracker_Server {
    public partial class PCTrackerServerForm : Form {

        private TcpListener serverSocket = new TcpListener(IPAddress.Any, 8888);
        private Thread AcceptClients;
        private List<Site> siteList = new List<Site>();
        //private string FilePath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\PC Tracker\\";
        private string FilePath = "C:\\Users\\Tech\\PC Tracker\\";
        private string SiteFileName = "Sites.xlsx";
        private LoadingProgress ProgressBarForm;
        private int ProgressMax;
        private bool Changed;
        private bool WindowDrawn;
        private List<Client> ClientList = new List<Client>();

        private delegate void StringParameterDelegate(string value);

        public PCTrackerServerForm() {
            InitializeComponent();
        }

        /// <summary>
        ///     update the tbLog.Text with any server messages, from any and all threads
        /// </summary>
        /// <param name="message"></param>
        public void UpdateStatus(string message) {
            if (InvokeRequired) {
                // We're not in the UI thread, so we need to call BeginInvoke
                BeginInvoke(new StringParameterDelegate(UpdateStatus), new object[] { message });
                return;
            }
            // Must be on the UI thread if we've got this far
            if (!tbLog.IsDisposed) {
                tbLog.AppendText(message);
                tbLog.AppendText(Environment.NewLine);
            }
        }

        public void RemoveClient(Client client) {
            ClientList.Remove(client);
        }

        /// <summary>
        ///     start loading up after the windows draws on the desktop
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void PCTrackerServerForm_Activated(object sender, EventArgs e) {
            if (!WindowDrawn) {
                WindowDrawn = true;
                UpdateStatus(">>>> Starting Server <<<<");
                UpdateStatus("> Loading Sites...");
                bgwLoadSites.RunWorkerAsync();
                ProgressBarForm = new LoadingProgress("Loading Sites List");
                ProgressBarForm.ShowDialog();
            }
        }

        /// <summary>
        ///     open and read the sites from the requested site file
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bgwLoadSites_DoWork(object sender, DoWorkEventArgs e) {
            Excel.Application excelApp = new Excel.Application() {
                Visible = false,
                DisplayAlerts = false
            };
            Excel.Workbook workbook;
            Excel.Worksheet worksheet;
            
            workbook = excelApp.Workbooks.Open(FilePath + SiteFileName);
            worksheet = workbook.Worksheets.Item[1];

            int localSitesNum = (int)worksheet.Cells[1, 2].Value;

            FillSiteList(localSitesNum, worksheet);
            workbook.Close();
            excelApp.Quit();
        }

        /// <summary>
        ///     load the sites into the siteList variable
        /// </summary>
        /// <param name="sitesNum"></param>
        /// <param name="worksheet"></param>
        private void FillSiteList(int sitesNum, Excel.Worksheet worksheet) {
            ProgressMax = sitesNum;
            string siteName;
            for (int i = 1; i <= sitesNum; i++) {
                siteName = ((string)worksheet.Cells[i, 1].Value).Split(' ')[0];
                siteList.Add(new Site(siteName));
                UpdateStatus("> " + siteName);
                bgwLoadSites.ReportProgress(i);
            }
        }

        /// <summary>
        ///     update the progress bar
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bgwLoadSites_ProgressChanged(object sender, ProgressChangedEventArgs e) {
            if (ProgressBarForm.getProgressMaximum() != ProgressMax) {
                ProgressBarForm.setProgressMaximum(ProgressMax);
            }
            ProgressBarForm.updateProgress(e.ProgressPercentage);
        }

        /// <summary>
        ///     after filling the siteList variable, this will go on to load up the PCs from 
        ///     each site, both:
        ///     Hotswaps (temporary shell replacement) and
        ///     Loaners (temporary use PC while permanent PC is being worked on)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bgwLoadSites_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e) {
            ProgressBarForm.Close();
            //load up the PC's for each site
            foreach (Site site in siteList) {
                UpdateStatus("> Loading PCs for " + site.Name + "...");
                //sends the current site object to the backgroundworker as an argument
                bgwLoadPCs.RunWorkerAsync(site);
                ProgressBarForm = new LoadingProgress("Loading PC Lists");
                ProgressBarForm.ShowDialog();
            }
            //this begins the Asynchronous thread for the auto-save feature.
            bgwAutoSave.RunWorkerAsync();
            //this goes on to open the server socket for on-coming user connections.
            openConnection();
        }
        
        /// <summary>
        ///     
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bgwLoadPCs_DoWork(object sender, DoWorkEventArgs e) {
            Excel.Application excelApp = new Excel.Application() {
                Visible = false,
                DisplayAlerts = false
            };

            //interprets the argument sent to the backgroundworker as a Site object
            Site site = (Site)e.Argument;

            Excel.Workbook workbook = excelApp.Workbooks.Open(FilePath + site.Name + "\\Loaners.xlsx");
            addLaptopstoSite(false, site, workbook);

            workbook = excelApp.Workbooks.Open(FilePath + site.Name + "\\Hotswaps.xlsx");
            addLaptopstoSite(true, site, workbook);

            excelApp.Quit();
        }

        /// <summary>
        ///     opens the worksheet and gets the laptop information to be added to either
        ///     the site's hotswap or loaner list
        /// </summary>
        /// <param name="hotswaps"></param>
        /// <param name="site"></param>
        /// <param name="workbook"></param>
        private void addLaptopstoSite(bool hotswaps, Site site, Excel.Workbook workbook) {
            Excel.Worksheet currentSheet = workbook.Worksheets.Item[1];

            int lastRow = getMaxRow(currentSheet);
            ProgressMax = lastRow;

            Laptop newLaptop;
            Laptop prevLaptop = new Laptop();

            for (int index = 2; index <= lastRow; index++) {
                newLaptop = getNewLaptop(index, ref currentSheet);
                //this verifies that the newly created laptop is not a copy of the previous one
                if (newLaptop != prevLaptop) {
                    if (hotswaps) {
                        site.Hotswaps.Add(newLaptop);
                    } else {
                        site.Loaners.Add(newLaptop);
                    }
                    UpdateStatus(newLaptop.Brand + " " + newLaptop.Model + " " + newLaptop.Serial);
                    bgwLoadPCs.ReportProgress(index);
                    prevLaptop = newLaptop;
                }
            }
            workbook.Close();
        }

        /// <summary>
        ///     reads each column of the indexed row to create a laptop object
        /// </summary>
        /// <param name="index"></param>
        /// <param name="sheet"></param>
        /// <returns>
        ///     a new Laptop object with information read from the sheet
        /// </returns>
        private Laptop getNewLaptop(int index, ref Excel.Worksheet sheet) {
            int lastCol = getMaxCol(sheet);
            // this array holds all of the information from each line of the excel sheet
            Array laptopValues = (Array)sheet.get_Range("A" + index.ToString(), ColumnNumToString(lastCol) + index.ToString()).Cells.Value;
            // I have to run the check null on each of these parsed cells, 
            // due to being brought in from an excel sheet with possible blank cells
            Laptop newLaptop = new Laptop() {
                Number = intCheckNull(laptopValues.GetValue(1, 1)),
                Serial = stringCheckNull(laptopValues.GetValue(1, 2)),
                Brand = stringCheckNull(laptopValues.GetValue(1, 3)),
                Model = stringCheckNull(laptopValues.GetValue(1, 4)),
                Warranty = stringCheckNull(laptopValues.GetValue(1, 5)),
                Username = stringCheckNull(laptopValues.GetValue(1, 6)),
                UserPCSerial = stringCheckNull(laptopValues.GetValue(1, 7)),
                TicketNumber = stringCheckNull(laptopValues.GetValue(1, 8)),
                CheckedOut = boolCheckNull(laptopValues.GetValue(1, 9))
            };
            return newLaptop;
        }

        /// <summary>
        ///     updates the progress bar of the dialog
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bgwLoadPCs_ProgressChanged(object sender, ProgressChangedEventArgs e) {
            if (ProgressBarForm.getProgressMaximum() != ProgressMax) {
                ProgressBarForm.setProgressMaximum(ProgressMax);
            }
            ProgressBarForm.updateProgress(e.ProgressPercentage);
        }

        /// <summary>
        ///     closes the progress bar dialog, once completed.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bgwLoadPCs_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e) {
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
            if (value == null) {
                return 0;
            }
            //the following will parse the value out to make sure it is a number, before assigning and returning
            int parsedNum;
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
                dividend = ((dividend - modulo) / 26);
            }
            return strColumnName;
        }

        /// <summary>
        ///     Starts the server socket and creates a thread to continually accept new connections
        /// </summary>
        private void openConnection() {
            serverSocket.Start();
            UpdateStatus(">> Server Started");

            AcceptClients = new Thread(ConnectClient);
            AcceptClients.Start(serverSocket);
        }

        /// <summary>
        ///     awaits the connection of a client via the specified socket, and 
        ///     creates a new Client object thread
        /// </summary>
        /// <param name="parameter"></param>
        private void ConnectClient(object parameter) {
            var serverSocket = parameter as TcpListener;
            var clientSocket = new TcpClient();
            //create a permanent loop to accept clients, while the server is active
            while (true) {
                try {
                    clientSocket = serverSocket.AcceptTcpClient();
                    //create the client object, once a connection is accepted
                    Client newClient = new Client(clientSocket, this);
                    //add that new Client to the List of Client objects
                    ClientList.Add(newClient);
                    UpdateStatus(">> Client " + newClient.UserName + " connected!");
                    //send the sites to the client, after connection is established
                    SendSitesToClient(newClient);
                } catch (Exception ex) {
                    UpdateStatus("XXX " + ex.Message);
                    break;
                }
            }
        }

        /// <summary>
        ///     gathers the site list, serializes the data and sends it to the client
        /// </summary>
        /// <param name="client"></param>
        private void SendSitesToClient(Client client) {
            UpdateStatus(">>> Sending sites to: " + client.UserName);
            
            //create a jagged array to store each serialzed site name
            byte[][] serializedData = new byte[siteList.Count][];
            foreach(Site site in siteList) {
                int index = siteList.IndexOf(site);
                serializedData[index] = new byte[site.Name.Length];
                serializedData[index] = SerializeString(site.Name);
            }

            List<byte> fullDataStream = new List<byte>();
            //add the serialized data for each site to a List to be sent to the client
            foreach (byte[] array in serializedData) {
                fullDataStream.AddRange(array);
            }
            client.StreamDataToClient(fullDataStream.ToArray(), this);
        }

        /// <summary>
        ///     serialize a string for data transfer
        /// </summary>
        /// <param name="s"></param>
        /// <returns>
        ///     a byte array with the data of the string
        /// </returns>
        private byte[] SerializeString(string s) {
            string stringToSerialize = s.Insert(s.Length, ";");
            byte[] serializedString = Encoding.UTF8.GetBytes(stringToSerialize);
            return serializedString;
        }
        
        /// <summary>
        ///     gathers the data of all the requested PCs of the requested type and
        ///     sends them to the client
        /// </summary>
        /// <param name="client"></param>
        /// <param name="siteName"></param>
        /// <param name="type"></param>
        public void SendPCsForSite(Client client, string siteName, string type) {
            var dataStream = new List<byte>();
            //add the identifier to the data stream
            dataStream.AddRange(BitConverter.GetBytes((int)DataIdentifier.Laptop));
            dataStream.AddRange(BitConverter.GetBytes(';'));
            //get the requested site from the List
            Site site = siteList.Find(s => s.Name == siteName);
            if (type == "Hotswaps") {
                foreach (Laptop pc in site.Hotswaps) {
                    dataStream.AddRange(pc.SerializeLaptop());
                }
            } else {
                foreach (Laptop pc in site.Loaners) {
                    dataStream.AddRange(pc.SerializeLaptop());
                }
            }
            UpdateStatus(">>> Sending " + type + " from " + siteName + " to " + client.UserName);
            client.StreamDataToClient(dataStream.ToArray(), this);
        }

        /// <summary>
        ///     
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void saveToolStripMenuItem_Click(object sender, EventArgs e) {
            SaveChanges();
        }

        /// <summary>
        ///     this will iterate through each site and send that site's Lists for 
        ///     saving in the spreadsheets
        /// </summary>
        /// <param name="siteName"></param>
        /// <param name="hotswaps"></param>
        private void SaveChanges() {
            Excel.Application excelApp = new Excel.Application() {
                Visible = false,
                DisplayAlerts = false
            };

            foreach (Site site in siteList) {
                UpdateStatus("<< Saving " + site.Name + "'s PC lists");
                Excel.Workbook workbook = excelApp.Workbooks.Open(FilePath + site.Name + "\\Loaners.xlsx");

                FillSheet(site.Loaners, workbook);

                workbook = excelApp.Workbooks.Open(FilePath + site.Name + "\\Hotswaps.xlsx");

                FillSheet(site.Hotswaps, workbook);
            }
            excelApp.Quit();
            Changed = false;
        }

        /// <summary>
        ///     read each part of each laptop and place it in the spreadsheets
        /// </summary>
        /// <param name="PCs"></param>
        /// <param name="workbook"></param>
        private void FillSheet(List<Laptop> PCs, Excel.Workbook workbook) {
            Excel.Worksheet sheet = workbook.Worksheets.Item[1];
            int lastrow = 2;
            foreach (Laptop PC in PCs) {
                sheet.Rows[lastrow].Delete();
                sheet.Cells[lastrow, 1].Value = PC.Number.ToString();
                sheet.Cells[lastrow, 2].Value = PC.Serial;
                sheet.Cells[lastrow, 3].Value = PC.Brand;
                sheet.Cells[lastrow, 4].Value = PC.Model;
                sheet.Cells[lastrow, 5].Value = PC.Warranty;
                sheet.Cells[lastrow, 6].Value = PC.Username;
                sheet.Cells[lastrow, 7].Value = PC.UserPCSerial;
                sheet.Cells[lastrow, 8].Value = PC.TicketNumber;
                sheet.Cells[lastrow, 9].Value = PC.CheckedOut;
                lastrow++;
            }
            workbook.Save();
            workbook.Close();
        }

        /// <summary>
        ///     runs the save function at set intervals
        ///     You'll notice a lot of things hinge on whether or not there is a cancellation
        ///     pending, this is so when the user turns the auto-save feature off, it will 
        ///     just move it's way out of the method without running the save feature again.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bgwAutoSave_DoWork(object sender, DoWorkEventArgs e) {
            UpdateStatus("** AutoSave Enabled **");
            while (!bgwAutoSave.CancellationPending) {
                DateTime start = DateTime.Now;
                //TODO: add the ability for the user to set the interval
                while (DateTime.Now.Subtract(start).Minutes < 30) {
                    if (bgwAutoSave.CancellationPending) {
                        break;
                    }
                }
                if (!bgwAutoSave.CancellationPending) {
                    SaveChanges();
                    UpdateStatus("<< Saving completed!");
                    UpdateStatus("<< Saving Log!");
                    string date = DateTime.Now.Year.ToString() + "-" + DateTime.Now.Day.ToString() + "-"
                        + DateTime.Now.Month.ToString() + " " + DateTime.Now.ToString("HH.mm.ss tt");
                    string logFile = FilePath + "logs\\log - " + date + ".txt";
                    if (!File.Exists(logFile)) {
                        File.Create(logFile).Dispose();
                    }
                    File.AppendAllText(logFile, tbLog.Text);
                }
            }
        }
        
        /// <summary>
        ///     
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bgwAutoSave_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e) {
            UpdateStatus("** Warning: AutoSave Disabled **");
        }

        /// <summary>
        ///     
        /// </summary>
        /// <param name="changedPC"></param>
        /// <param name="client"></param>
        public void updatePC(PCChange changedPC, Client client) {
            Site site = siteList.Find(s => s.Name == client.Site);
            Laptop PCtoEdit;

            if (client.Hotswaps) {
                PCtoEdit = site.Hotswaps.Find(pc => pc.Serial == changedPC.Serial);
            } else {
                PCtoEdit = site.Loaners.Find(pc => pc.Serial == changedPC.Serial);
            }

            string modification = string.Empty;
            if (changedPC.CheckedOut) {
                modification = " is checking out ";
            } else {
                modification = " is checking in ";
            }
            UpdateStatus(">>> User " + client.UserName + modification + PCtoEdit.Serial +
                         " from site " + client.Site);
            Changed = true;
            PCtoEdit.MergeChanges(changedPC);

            UpdateStatus(">> Sending updates to other clients connected to " + client.Site);
            BroadcastUpdateToSite(PCtoEdit, client);
        }

        /// <summary>
        ///     sends out the received message to each client that is connected.
        /// </summary>
        /// <param name="packet"></param>
        /// <param name="flag"></param>
        public void BroadcastUpdateToSite(string broadcastMsg, Client client) {
            var serializedData = new List<byte>();
            serializedData.AddRange(BitConverter.GetBytes((int)DataIdentifier.Broadcast));
            serializedData.AddRange(SerializeString(broadcastMsg));
            foreach (Client c in ClientList.FindAll(c => (c.Site == client.Site && c.Hotswaps == client.Hotswaps))) {
                if(c != client) {
                    UpdateStatus(">>> Sending update broadcast to " + c.UserName);
                    c.StreamDataToClient(serializedData.ToArray(), this);
                }
            }
        }

        /// <summary>
        ///     sends out the broadcast to each client that is connected.
        /// </summary>
        /// <param name="packet"></param>
        /// <param name="flag"></param>
        public void BroadcastUpdateToSite(Laptop updatedPC, Client client) {
            var serializedData = new List<byte>();
            serializedData.AddRange(BitConverter.GetBytes((int)DataIdentifier.Update));
            serializedData.AddRange(updatedPC.SerializeLaptop());
            foreach (Client c in ClientList.FindAll(c => (c.Site == client.Site && c.Hotswaps == client.Hotswaps))) {
                if (c != client) {
                    UpdateStatus(">>> Sending update broadcast to " + c.UserName);
                    c.StreamDataToClient(serializedData.ToArray(), this);
                }
            }
        }

        /// <summary>
        ///     
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void quitToolStripMenuItem_Click(object sender, EventArgs e) {
            Close();
        }

        /// <summary>
        ///     
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void PCTrackerServerForm_Closing(object sender, FormClosingEventArgs e) {
            serverSocket.Stop();
            AcceptClients.Abort();
            do {
                bgwAutoSave.CancelAsync();
            } while (!bgwAutoSave.IsBusy);
            if (Changed) {
                SaveChanges();
            }
            string date = DateTime.Now.Year.ToString() + "-" + DateTime.Now.Day.ToString() + "-" + DateTime.Now.Month.ToString();
            string logFile = FilePath + "logs\\log - " + date + ".txt";
            if (!File.Exists(logFile)) {
                File.Create(logFile).Dispose();
            }
            File.AppendAllText(logFile, tbLog.Text);
        }

        /// <summary>
        ///     
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void testBroadcastToolStripMenuItem_Click(object sender, EventArgs e) {
            string test = "abcdefghijklmnopqrstuvwxyz1234567890-=[]\\',./`ABCDEFGHIJKLMNOPQRSTUVWXYZ!@#$%^&*()_+{}|:\"<>?~";
            var serializedData = new List<byte>();
            serializedData.AddRange(BitConverter.GetBytes((int)DataIdentifier.Broadcast));
            serializedData.AddRange(SerializeString(test));
            foreach (Client c in ClientList) {
                UpdateStatus(">> Sending Test Broadcast to " + c.UserName);
                c.StreamDataToClient(serializedData.ToArray(), this);
            }
        }

        /// <summary>
        ///     
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void autoSaveToolStripMenuItem_Click(object sender, EventArgs e) {
            if (autoSaveToolStripMenuItem.Checked) {
                if (!bgwAutoSave.IsBusy) {
                    bgwAutoSave.RunWorkerAsync();
                }
            } else {
                bgwAutoSave.CancelAsync();
            }
        }
    }
}
