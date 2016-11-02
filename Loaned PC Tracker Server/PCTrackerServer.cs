using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Net;
using System.Net.Sockets;

namespace Loaned_PC_Tracker_Server {
    public partial class PCTrackerServerForm : Form {

        private delegate void StringParameterDelegate(string value);
        private TcpListener serverSocket = new TcpListener(IPAddress.Any, 8888);
        private Thread AcceptClients;
        private List<Site> siteList = new List<Site>();
        private Excel.Application excelApp = new Excel.Application() {
            Visible = false,
            DisplayAlerts = false
        };
        private string FilePath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\PC Tracker\\";
        private string SiteFileName = "Sites.xlsx";
        private LoadingProgress ProgressBarForm;
        //private int Counter = 0;
        private int ProgressMax;
        private bool Changed;
        private bool WindowDrawn;

        public List<Client> ClientList = new List<Client>();

        public PCTrackerServerForm() {
            InitializeComponent();
        }

        /// <summary>
        ///     
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void PCTrackerServerForm_Activated(object sender, EventArgs e) {
            if (!WindowDrawn) {
                WindowDrawn = true;
                tbLog.AppendText("Loading Sites...");
                tbLog.AppendText(Environment.NewLine);
                bgwLoadSites.RunWorkerAsync();
                ProgressBarForm = new LoadingProgress("Loading Sites List");
                ProgressBarForm.ShowDialog();
            }
        }

        /// <summary>
        ///     
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bgwLoadSites_DoWork(object sender, DoWorkEventArgs e) {
            Excel.Workbook workbook;
            Excel.Worksheet worksheet;
            
            workbook = excelApp.Workbooks.Open(FilePath + SiteFileName);
            worksheet = workbook.Worksheets.Item[1];

            int localSitesNum = (int)worksheet.Cells[1, 2].Value;

            FillSiteList(localSitesNum, worksheet);
            workbook.Close();
        }

        /// <summary>
        ///     
        /// </summary>
        /// <param name="sitesNum"></param>
        /// <param name="worksheet"></param>
        private void FillSiteList(int sitesNum, Excel.Worksheet worksheet) {
            ProgressMax = sitesNum;
            string siteName;
            for (int i = 1; i <= sitesNum; i++) {
                siteName = ((string)worksheet.Cells[i, 1].Value).Split(' ')[0];
                siteList.Add(new Site(siteName));
                bgwLoadSites.ReportProgress(i, siteName);
            }
        }

        /// <summary>
        ///     
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bgwLoadSites_ProgressChanged(object sender, ProgressChangedEventArgs e) {
            if (ProgressBarForm.getProgressMaximum() != ProgressMax) {
                ProgressBarForm.setProgressMaximum(ProgressMax);
            }
            ProgressBarForm.updateProgress(e.ProgressPercentage);
            tbLog.AppendText((string)e.UserState);
            tbLog.AppendText(Environment.NewLine);
        }

        /// <summary>
        ///     
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bgwLoadSites_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e) {
            ProgressBarForm.Close();
            foreach (Site site in siteList) {
                tbLog.AppendText("Loading PCs for " + site.Name + "...");
                tbLog.AppendText(Environment.NewLine);
                bgwLoadPCs.RunWorkerAsync(site);
                ProgressBarForm = new LoadingProgress("Loading PC Lists");
                ProgressBarForm.ShowDialog();
            }
            openConnection();
        }
        
        /// <summary>
        ///     
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bgwLoadPCs_DoWork(object sender, DoWorkEventArgs e) {
            //string fileName = (string)e.Argument;
            Site site = (Site)e.Argument;
            Excel.Workbook workbook = excelApp.Workbooks.Open(FilePath + site.Name + "\\Loaners.xlsx");
            Excel.Worksheet currentSheet = workbook.Worksheets.Item[1];

            int lastRow = getMaxRow(currentSheet);
            ProgressMax = lastRow;

            Laptop newLaptop;
            Laptop prevLaptop = new Laptop();

            for (int index = 2; index <= lastRow; index++) {
                newLaptop = getNewLaptop(index, ref currentSheet);
                //this verifies that the newly created laptop is not a copy of the previous one
                if (newLaptop != prevLaptop) {
                    if (newLaptop.CheckedOut) {
                        site.CheckedOutLoaners.Add(newLaptop);
                    } else {
                        site.AvailableLoaners.Add(newLaptop);
                    }
                    bgwLoadPCs.ReportProgress(index, newLaptop.Serial);
                    prevLaptop = newLaptop;
                }
            }
            workbook.Close();

            workbook = excelApp.Workbooks.Open(FilePath + site.Name + "\\Hotswaps.xlsx");
            currentSheet = workbook.Worksheets.Item[1];

            lastRow = getMaxRow(currentSheet);
            ProgressMax = lastRow;

            prevLaptop = new Laptop();

            for (int index = 2; index <= lastRow; index++) {
                newLaptop = getNewLaptop(index, ref currentSheet);
                //this verifies that the newly created laptop is not a copy of the previous one
                if (newLaptop != prevLaptop) {
                    if (newLaptop.CheckedOut) {
                        site.CheckedOutHotswaps.Add(newLaptop);
                    } else {
                        site.AvailableHotswaps.Add(newLaptop);
                    }
                    bgwLoadPCs.ReportProgress(index, newLaptop.Serial);
                    prevLaptop = newLaptop;
                }
            }
            workbook.Close();
        }

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
        ///     
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bgwLoadPCs_ProgressChanged(object sender, ProgressChangedEventArgs e) {
            tbLog.AppendText((string)e.UserState);
            tbLog.AppendText(Environment.NewLine);

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
                dividend = ((dividend - modulo) / 26);
            }
            return strColumnName;
        }

        /// <summary>
        ///     
        /// </summary>
        private void openConnection() {
            serverSocket.Start();
            tbLog.AppendText(" >> Server Started");
            tbLog.AppendText(Environment.NewLine);

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
            var clientSocket = default(TcpClient);
            while (true) {
                try {
                    clientSocket = serverSocket.AcceptTcpClient();
                    Client newClient = new Client(clientSocket);
                    //newClient.startClient(clientSocket);
                    ClientList.Add(newClient);
                    UpdateStatus("Client " + newClient.UserName + " connected!");
                    SendSitesToClient(newClient);
                } catch (SocketException ex) {
                    //UpdateStatus(" >> Something Happened!! ");
                    //UpdateStatus(ex.Message.ToString());
                    //clientSocket.Close();
                    break;
                }
            }
        }

        /// <summary>
        ///     
        /// </summary>
        /// <param name="message"></param>
        private void UpdateStatus(string message) {
            if (InvokeRequired) {
                // We're not in the UI thread, so we need to call BeginInvoke
                BeginInvoke(new StringParameterDelegate(UpdateStatus), new object[] { message });
                return;
            }
            // Must be on the UI thread if we've got this far
            tbLog.AppendText(message);
            tbLog.AppendText(Environment.NewLine);
        }

        private void SendSitesToClient(Client client) {
            NumberPacket numSites = new NumberPacket(siteList.Count);
            client.SendPacketToClient(numSites);
            foreach(Site site in siteList) {
                client.SendPacketToClient(new NamePacket(site.Name));
            }
        }

        /// <summary>
        ///     sends out the broadcasted chat or system message to each client that is connected.
        /// </summary>
        /// <param name="packet"></param>
        /// <param name="flag"></param>
        public void Broadcast(PCPacket packet, bool flag = true) {
            foreach (Client client in ClientList) {
                client.SendPacketToClient(packet);
            }
        }

        /// <summary>
        ///     
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void saveToolStripMenuItem_Click(object sender, EventArgs e) {
            //SaveChanges();
        }

        /// <summary>
        ///     
        /// </summary>
        /// <param name="siteName"></param>
        /// <param name="hotswaps"></param>
        private void SaveChanges() {
            foreach (Site site in siteList) {
                tbLog.AppendText("Saving " + site.Name + "'s PC lists");
                tbLog.AppendText(Environment.NewLine);
                bgwSaveChanges.RunWorkerAsync(site);
                ProgressBarForm = new LoadingProgress("Saving PC Lists");
                ProgressBarForm.ShowDialog();
            }
        }

        /// <summary>
        ///     
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bgwSaveChanges_DoWork(object sender, DoWorkEventArgs e) {
            Site site = (Site)e.Argument;
            Excel.Workbook workbook = excelApp.Workbooks.Open(FilePath + site.Name + "\\Loaners.xlsx");
            Excel.Worksheet currentSheet = workbook.Worksheets.Item[1];

            int lastrow = 2;
            int progress = 0;
            ProgressMax = (site.AvailableLoaners.Count + site.CheckedOutLoaners.Count);

            //todo: add logic to replace the data in the sheet with the new data
            foreach (Laptop PC in site.AvailableLoaners.Union(site.CheckedOutLoaners)) {
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
                bgwSaveChanges.ReportProgress(++progress, PC.Serial);
            }
            workbook.Save();
            workbook.Close();

            workbook = excelApp.Workbooks.Open(FilePath + site.Name + "\\Hotswaps.xlsx");
            currentSheet = workbook.Worksheets.Item[1];

            lastrow = 2;
            progress = 0;
            ProgressMax = (site.AvailableHotswaps.Count + site.CheckedOutHotswaps.Count);

            //todo: add logic to replace the data in the sheet with the new data
            foreach (Laptop PC in site.AvailableHotswaps.Union(site.CheckedOutHotswaps)) {
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
                bgwSaveChanges.ReportProgress(++progress, PC.Serial);
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
            tbLog.AppendText((string)e.UserState);
            tbLog.AppendText(Environment.NewLine);
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
            Changed = false;
            ProgressBarForm.Close();
        }

        /// <summary>
        ///     
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void quitToolStripMenuItem_Click(object sender, EventArgs e) {
            if (!Changed) {
                Close();
            }
        }

        /// <summary>
        ///     
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void PCTrackerServerForm_Closing(object sender, FormClosingEventArgs e) {
            excelApp.Quit();
            //AcceptClients.Interrupt();
            serverSocket.Stop();
        }
    }
}
