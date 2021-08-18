using System;
using System.Configuration;
using System.Windows.Forms;
using TransformExcel;
using SICommon;
using System.IO;
using System.Data;
using System.Data.SqlClient;
using System.Text;
using LineItemCancels;

namespace FixContractItems
{
    public partial class frmFixContractItems : Form
    {
        private ProcessExcel _processExcel;
        log4net.ILog log;
        string sConnectionString;
        Int32 maxToRead = 0;
        Int32 maxToWrite = 0;
        string Tier = "";
        string tempDir;
        string Version = ".9";
        DataTable dtMismatches;
        LineItems _lineItem;

        public frmFixContractItems()
        {
            InitializeComponent();

            SetUp();
        }

        private void SetUp()
        {
            log4net.Config.XmlConfigurator.Configure();
            log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

            log.Info("Starting frmFixContractItems version:" + Version);
            sConnectionString = ConfigurationManager.ConnectionStrings["ServerDB"].ToString();
            Tier = SIAppRoutines.RetrieveParmString("Tier", "");
            maxToRead = SIAppRoutines.RetrieveParmInteger("maxToRead", 0);
            maxToWrite = SIAppRoutines.RetrieveParmInteger("maxToWrite", 0);

            tempDir = GetTempDir();

            log.Debug("TempDir:" + tempDir + " max to read:" + maxToRead.ToString() + " max to write:" + maxToWrite.ToString());
            log.Debug("ConnectionString:" + sConnectionString);

            _lineItem = new LineItems();
            _lineItem.ConnectionStringSQL = sConnectionString;

        }

        private void mnuExit_Click(object sender, EventArgs e)
        {
            ShutDown();
        }
        private void ShutDown()
        {
            Application.Exit();
        }
        private void ExpandDataGridView(Button vButton, DataGridView vDGV)
        {
            if (vButton.Text == "+")
            {
                vButton.Text = "-";
                vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            }
            else
            {
                vButton.Text = "+";
                vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCellsExceptHeader;
            }

            vDGV.Refresh();
        }
        private string GetTempDir()
        {
            string sTempDir = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\FixContractItems";
            try
            {
                if (!Directory.Exists(sTempDir))
                    Directory.CreateDirectory(sTempDir);
            }
            catch (UnauthorizedAccessException ex)
            {
                MessageBox.Show(ex.Message + " Could not create a temporary directory on: " + sTempDir, ", Unable to Create Temp Files.", MessageBoxButtons.OK);
                return null;
            }

            return sTempDir;
        }
        private bool WriteExcelFile(TextBox vFileName, TextBox vSheetName, DataTable vdt, bool vIsWriteHeaders, ProcessExcel vProcessExcel=null)
        {
            Properties.Settings.Default.Save();

            string sOutputFileName = vFileName.Text.Trim();
            string sSheetName = vSheetName.Text.Trim();


            if (sOutputFileName.Length == 0)
            {
                MessageBox.Show("You must enter an output file name");
                return false;
            }

            if (sSheetName.Length == 0)
            {
                MessageBox.Show("You must enter an output sheet name");
                return false;
            }

            if (vProcessExcel == null)
            {
                vProcessExcel = new ProcessExcel();
            }

            if (vdt == null)
            {
                MessageBox.Show("Input Data Table Not Defined");
                return false;
            }

            vProcessExcel.isWriteHeaders = vIsWriteHeaders;
            vProcessExcel.WriteDataTableToExcelFile(vdt, sOutputFileName, sSheetName);
            if (vProcessExcel.isFatalError)
            {
                MessageBox.Show("Error:" + vProcessExcel.FatalErrorMessage);
                return false;
            }
            else
            {
                MessageBox.Show("Finished Writing File");
            }

            return true;

        }

        private string FindFileName(TextBox vTextBox)
        {
            string sLastPath = "";

            var OpenFileDialog = new OpenFileDialog();

            if (OpenFileDialog.ShowDialog(this) != DialogResult.OK)
            {
                return "";
            }

            sLastPath = OpenFileDialog.FileName;
            vTextBox.Text = sLastPath;
            return sLastPath;
        }

        private void btnMismatchedExpand_Click(object sender, EventArgs e)
        {
            ExpandDataGridView((Button)sender, dgvMismatched);

        }

        private void btnWriteMismatched_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            Button btn = (Button)sender;
            btn.Enabled = false;

            bool isWriteHeaders = true;

            WriteExcelFile(txtMismatchedOutputFileName, txtMismatchedSheetName, dtMismatches, isWriteHeaders);

            Cursor = Cursors.Default;
            btn.Enabled = true;

        }

        private void btnRetrieveMismatched_Click(object sender, EventArgs e)
        {
            Button btn = (Button)sender;
            btn.Enabled = false;
            Cursor = Cursors.WaitCursor;

            dtMismatches = retrieveMismatchedItems();

            lblMismatchedRecordCount.Text = dtMismatches.Rows.Count.ToString();

            AddWorkColumns();

            dgvMismatched.DataSource = dtMismatches;
            dgvMismatched.Refresh();

            btn.Enabled = true;
            Cursor = Cursors.Default;

        }

        private DataTable retrieveMismatchedItems()
        {
            DataTable dt = new DataTable();

            StringBuilder sbSQL = new StringBuilder();

            sbSQL.Append("select ");
            sbSQL.Append("contractitems.uniquekey as contractitemsUK, ");
            sbSQL.Append("contractitems.lastupdate as ciupdate, ");
            sbSQL.Append("batchcontract.InvoiceNo, ");
            sbSQL.Append("batchcontract.batchnum as bcbatchnum, ");
            sbSQL.Append("contractitems.BatchNum as cibatchnum, ");
            sbSQL.Append("batchcontract.uniquekey as bcuk, ");
            sbSQL.Append("contractitems.BatchConNum as cibatchconnum, ");
            sbSQL.Append("cibatchcontract.uniquekey as batchcontractforCIuk, ");
            sbSQL.Append("batchcontract.Claim, ");
            sbSQL.Append("cibatchcontract.Claim as CIBCClaim, ");
            sbSQL.Append("batchcontract.entrydate, ");
            sbSQL.Append("cibatchcontract.entrydate as cibcEntrydate, ");
            sbSQL.Append("contractitems.description, ");
            sbSQL.Append("dealer.dealernumber as bcDealer, ");
            sbSQL.Append("dealer.dealername as bcDealerName, ");
            sbSQL.Append("dealer2.DealerNumber as ciDealer, ");
            sbSQL.Append("dealer2.dealername as ciDealerName, ");
            sbSQL.Append(" ");
            sbSQL.Append("batchcontract.cancelentrydate as bcCancelEntrydate, ");
            sbSQL.Append("cibatchcontract.cancelentrydate as ciBCCancelEntryDate, ");
            sbSQL.Append("contractitems.LastUpdate as ciLastupdate, ");
            sbSQL.Append("iif (batchcontract.sUniTran <> contractitems.sUniTran,'**','') as diffunitran, ");
            sbSQL.Append("batchcontract.sUniTran as bcunitran, ");
            sbSQL.Append("contractitems.sUniTran as ciunitran, ");
            sbSQL.Append("contractitems.* ");
            sbSQL.Append(" ");
            //sbSQL.Append("contractitems.*, ");
            //sbSQL.Append("batchcontract.* ");
            sbSQL.Append("from ");
            sbSQL.Append("contractitems ");
            sbSQL.Append("inner join batchcontract on batchcontract.uniquekey = contractitems.BatchConNum ");
            sbSQL.Append("inner join batchheader on batchheader.uniquekey = batchcontract.BatchNum ");
            sbSQL.Append("left join batchheader bh2 on bh2.uniquekey = contractitems.BatchNum ");
            sbSQL.Append("inner join dealer on dealer.uniquekey = batchheader.dealer ");
            sbSQL.Append("left join dealer dealer2 on dealer2.uniquekey = bh2.dealer ");
            sbSQL.Append("inner join product on product.uniquekey = batchcontract.iProduct ");
            sbSQL.Append("left join batchcontract cibatchcontract on cibatchcontract.sUniTran = contractitems.sUniTran ");
            sbSQL.Append("where ");
            sbSQL.Append("contractitems.BatchNum <> batchcontract.BatchNum ");
            sbSQL.Append("and convert(date,batchcontract.entrydate) >= '2021-01-25' ");
            sbSQL.Append("order by batchcontract.entrydate ");

            dt = RetrieveDataTable(sbSQL.ToString());

            return dt;
        }

        private DataTable RetrieveDataTable(string vSQL)
        {
            DataTable dt = new DataTable();

            try
            {
                using (SqlConnection oConn = new SqlConnection(sConnectionString))
                {
                    using (SqlDataAdapter oAdapter = new SqlDataAdapter())
                    {
                        try
                        {
                            oAdapter.SelectCommand = new SqlCommand(vSQL, oConn);
                            oAdapter.Fill(dt);
                        }
                        catch (Exception ex)
                        {
                            Cursor = Cursors.Default;
                            MessageBox.Show("SQL Error: " + ex.Message);
                            return null;

                        }
                    }
                }

            }
            catch (Exception ex2)
            {
                Cursor = Cursors.Default;
                MessageBox.Show("SQL Error: " + ex2.Message);
                return null;

            }

            return dt;

        }
        private void AddWorkColumns()
        {
            string[] columns = retrieveWorkColumns();
            ColumnUtilities.addColumnsToTable(dtMismatches, columns);

            ColumnUtilities.reorderColumns(dtMismatches, columns);

        }
        public string[] retrieveWorkColumns()
        {
            string[] columns = new[]
                    {
                        "isDestinationFound",
                        "fileDescription",

                        "destCount",
                        "Errors"
                    };

            return columns;

        }

        private void btnValidate_Click(object sender, EventArgs e)
        {
            Button btn = (Button)sender;
            btn.Enabled = false;
            Cursor = Cursors.WaitCursor;

            lblNotFoundCount.Text = "0";
            lblFoundCount.Text = "0";

            ValidateItems();

            log.Info ("Items: " + readCount.ToString() + " Found:" + itemsFound.ToString() + " not found:" + itemsNotFound.ToString());

            btn.Enabled = true;
            Cursor = Cursors.Default;

        }

        Int32 readCount = 0;
        Int32 itemsFound = 0;
        Int32 itemsNotFound = 0;

        private void ValidateItems()
        {
            readCount = 0;
            itemsFound = 0;
            itemsNotFound = 0;

            foreach (DataRow myRow in dtMismatches.Rows)
            {
                if (maxToRead > 0 && readCount > maxToRead)
                {
                    break;
                }

                readCount += 1;

                lblStatus.Text = readCount.ToString();
                Application.DoEvents();

                RetrieveItemInfo(myRow);
                if (_lineItem.isFatalError)
                {
                    MessageBox.Show("Fatal Error: " + _lineItem.FatalErrorMessage);
                    return;
                }
            }
        }

        Int32 prevBatchConNum = -1;
        Int32 BatchConNum = 0;

        private void RetrieveItemInfo(DataRow vRow)
        {
            //vRow["isDestinationFound"] = "y";
            BatchConNum = (Int32) vRow["batchcontractforCIuk"];
            string fileDescription = vRow["description"] + " - ";

            if (BatchConNum != prevBatchConNum)
            {
                _lineItem.batchConNum = BatchConNum;
                _lineItem.isMatchItemID = false;
                _lineItem.isMatchModel = false;

                _lineItem.RetrieveContractItemsInfo();

                vRow["destCount"] = _lineItem.colContractItemsDataBase.Count.ToString();

                prevBatchConNum = BatchConNum;
                foreach (ContractItemsInput vInput in _lineItem.colContractItemsDataBase)
                {
                    //vInput.Model = "";  //space model if it is vastly different but not significant 
                    fileDescription += vInput.Description;
                }
            }

            vRow["fileDescription"] = fileDescription;
            vRow["destCount"] = _lineItem.colContractItemsDataBase.Count.ToString();

            ContractItemsInput _contractItem = new ContractItemsInput();

            _contractItem.Description = vRow["Description"].ToString();
            _contractItem.MFG = vRow["MFG"].ToString();
            _contractItem.BatchNum = (Int32)vRow["BatchNum"];
            _contractItem.BatchConNum = (Int32) vRow["batchConNum"];
            _contractItem.Model = vRow["Model"].ToString();
            //_contractItem.Model = "";    //blank model if it is different but not significant in finding the right contract
            _contractItem.SerialNum = vRow["SerialNum"].ToString();
            _contractItem.SerialNum = vRow["sItemID"].ToString();
            _contractItem.SerialNum = vRow["sUniTran"].ToString();
            _contractItem.Quantity = (Int32) vRow["Quantity"];
            _contractItem.Cost = (double)vRow["cost"];

            _lineItem.isMatchModel = false;
            _lineItem.isMatchItemID = false;
            bool isMatch = _lineItem.MatchItem(_contractItem, _lineItem.colContractItemsDataBase);
            if (isMatch)
            {
                itemsFound += 1;
                lblFoundCount.Text = itemsFound.ToString();
            }
            else
            {
                itemsNotFound += 1;
                lblNotFoundCount.Text = itemsNotFound.ToString();
            }

            vRow["isDestinationFound"] = isMatch.ToString();
        }

        private void btnWriteDeleteSQL_Click(object sender, EventArgs e)
        {
            string fileName = txtOutputFileName.Text.Trim();
 
            if (fileName.Length == 0)
            {
                MessageBox.Show("You must enter an output file name - Retry");
                return;
            }

            Properties.Settings.Default.Save();

            CreateDeleteSQLForMatchedItems(fileName);

            MessageBox.Show("Output file written");

        }

        private void CreateDeleteSQLForMatchedItems(string vFileName)
        {
            StreamWriter outputFile = new StreamWriter(vFileName);

            Int32 numFound = 0;
            Int32 numNotFound = 0;

            foreach (DataRow myRow in dtMismatches.Rows)
            {
                string tf = myRow["isDestinationFound"].ToString();
                string uk = myRow["contractitemsUK"].ToString();
                string badBatchNum = myRow["cibatchnum"].ToString();
                if (tf.ToLower() != "true")
                {
                    {
                        numNotFound += 1;

                        //outputFile.WriteLine("isDestinationfound Value:" + tf + " False:" + uk);
                        continue;
                    }
                }

                numFound += 1;

                string sSQLFmt = "Delete from contractitems where uniquekey = '" + uk.ToString() + "' and batchnum = '" + badBatchNum + "';";
                outputFile.WriteLine(sSQLFmt);

            }

            outputFile.WriteLine("--Num Found:" + numFound.ToString() + " not found:" + numNotFound.ToString());
            outputFile.Close();
           
        }

        private void btnWriteMoveSQL_Click(object sender, EventArgs e)
        {
            string fileName = txtOutputFileName.Text.Trim();

            if (fileName.Length == 0)
            {
                MessageBox.Show("You must enter an output file name - Retry");
                return;
            }

            Properties.Settings.Default.Save();

            CreateMoveSQLForUnMatchedItems(fileName);

            MessageBox.Show("Output file written");
        }

        private void CreateMoveSQLForUnMatchedItems(string vFileName)
        {
            StreamWriter outputFile = new StreamWriter(vFileName);

            Int32 numFound = 0;
            Int32 numNotFound = 0;

            foreach (DataRow myRow in dtMismatches.Rows)
            {
                string tf = myRow["isDestinationFound"].ToString();
                string uk = myRow["contractitemsUK"].ToString();
                string badBatchNum = myRow["cibatchnum"].ToString();
                string destCount = myRow["destCount"].ToString();
                string newCIBatchConNum = myRow["batchcontractforCIUk"].ToString();
                string badBatchConNum = myRow["cibatchconnum"].ToString();

                if (tf.ToLower() != "false")
                {
                    {
                        numNotFound += 1;

                        //outputFile.WriteLine("isDestinationfound Value:" + tf + " False:" + uk);
                        continue;
                    }
                }

                if (destCount != "0")
                {
                    continue;
                }

                numFound += 1;

                string sSQLFmt = "update  contractitems set batchconnum = '" + newCIBatchConNum + 
                    "' where uniquekey = '" + uk.ToString() + "' and batchnum = '" + badBatchNum + "' and batchconnum ='" + badBatchConNum + "';";
                outputFile.WriteLine(sSQLFmt);

            }

            outputFile.WriteLine("--Num Found:" + numFound.ToString() + " not found:" + numNotFound.ToString());
            outputFile.Close();

        }
    }
}
