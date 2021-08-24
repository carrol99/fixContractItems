﻿using System;
using System.Configuration;
using System.Windows.Forms;
using TransformExcel;
using SICommon;
using System.IO;
using System.Data;
using System.Data.SqlClient;
using System.Text;
using LineItemCancels;
using ContractExcelToXml;

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
        private string FatalErrorMessage;
        private bool isFatalError;
        private DbDataRoutines _SqlDataRoutines;
        DataTable dtInput;
        DataTable dtOutput;
        Int32 ContractsWithNoItems = 0;
        Int32 ContractsWithItems = 0;
        Int32 ContractsNotFound = 0;
        Int32 ContractsFound = 0;
        Int32 ItemsFound = 0;
        Int32 ItemsNotFound = 0;
        Int32 seqnum = 0;
        DateTime currentUpdateTime = DateTime.Now;
        Int32 skipRecords = 00;



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
            log.Debug("ConnectionString:" + SQLUtil.HideField( sConnectionString, "pwd"));

            _lineItem = new LineItems();
            _lineItem.ConnectionStringSQL = sConnectionString;

            _SqlDataRoutines = new DbDataRoutines(sConnectionString);

        }

        private void mnuExit_Click(object sender, EventArgs e)
        {
            ShutDown();
        }
        private void ShutDown()
        {
            Application.Exit();
        }

        #region "Misc Routines"
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

        private void LoadProcessInputPremiumFile(object sender, bool vIsLoadFileWithHeaders = false)
        {

            isFatalError = false;
            FatalErrorMessage = "";

            Properties.Settings.Default.Save();

            var btn = (Button)sender;
            btn.Enabled = false;
            Cursor = Cursors.WaitCursor;

            string InputExcelFileName = txtMissingFileName.Text.Trim();
            string InputSheetName = txtMissingSheetName.Text.Trim();

            ProcessExcel _process = new ProcessExcel();

            _process = RetrievePremiumExcelFile(InputExcelFileName, InputSheetName, vIsLoadFileWithHeaders);

            if (_process.isFatalError)
            {
                FatalErrorMessage = _process.FatalErrorMessage;
                isFatalError = true;
                MessageBox.Show("Error:" + FatalErrorMessage);
                btn.Enabled = true;
                Cursor = Cursors.Default;

                return;
            }

            //dtLookupProcess = _process.outputDataTable;
            dtInput = _process.inputDataTable;
            dtOutput = _process.outputDataTable;

            dgvProcessInput.DataSource = _process.outputDataTable;
            dgvProcessInput.Refresh();

            dgvProcessOutput.DataSource = _process.outputDataTable;
            dgvProcessOutput.Refresh();

            lblProcessOutputCount.Text = _process.outputDataTable.Rows.Count.ToString();

            btn.Enabled = true;
            Cursor = Cursors.Default;

        }

        public ProcessExcel RetrievePremiumExcelFile(string vFileName, string vSheetName, bool vLoadFileWithHeaders = false)
        {
            ProcessExcel _transformExcel = new TransformExcel.ProcessExcel();

            // attachTransFormEvents(_transformExcel);

            _transformExcel.isCleanUpHeaderColumns = true;

            _transformExcel.ExcelFileName = vFileName;
            _transformExcel.sheetName = vSheetName;
            _transformExcel.maxToRead = maxToRead;
            _transformExcel.SortOrder = "Auth_Number";

            if (vLoadFileWithHeaders)
            {
                _transformExcel.MapColumns = MapColumnDefinitions.buildColumnMappingsWarrantyPremiumWithHeadings();
                MapColumn map = new MapColumn
                {
                    inColumnName = "certificate",
                    isAlias = true,
                    inAliasNames =
                    {
                        "certificate_num",
                        "certno",
                        "cert",
                        "cert_#",
                        "cert#",
                        "certificate_number",
                        "certificate_no"
                    },
                    outColumnName = "certificate",
                    isOptional = true,
                    isPosition = false
                };
                _transformExcel.MapColumns.Add(map);
            }
            else
            {
                _transformExcel.MapColumns = MapColumnDefinitions.buildColumnMappingsWarrantyPremium();
            }

            _transformExcel.OutputColumnNames = _transformExcel.BuildOutputColumnsFromMapColumns();

            //_transformExcel.isTrimSpaces = _isTrimSpaces;

            _transformExcel.isNoHeader = true;

            string suffix = System.IO.Path.GetExtension(_transformExcel.ExcelFileName);

            if (suffix != null && suffix.ToLower() == ".xml")
            {
                ConvertToXml _contractXMLToExcel = new ConvertToXml();
                _contractXMLToExcel.XmlFileName = _transformExcel.ExcelFileName;
                _contractXMLToExcel.ConvertXMLToExcel();
                _transformExcel.outputDataTable = _contractXMLToExcel.dtExcelOutput;

            }
            else
            {
                _transformExcel.ConvertExcelFile();

                if (_transformExcel.isFatalError)
                {
                    isFatalError = true;
                    FatalErrorMessage = _transformExcel.FatalErrorMessage;
                    return _transformExcel;
                }
            }

            //dtExport = _transformExcel.outputDataTable.Clone();
            //dtExportProcess = _transformExcel.outputDataTable.Clone();

            return _transformExcel;
        }


        #endregion

        #region "Mismatched Items"
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

 
        private void btnValidate_Click(object sender, EventArgs e)
        {
            Button btn = (Button)sender;
            btn.Enabled = false;
            Cursor = Cursors.WaitCursor;

            lblNotFoundCount.Text = "0";
            lblFoundCount.Text = "0";

            ValidateItems();

            log.Info ("Items: " + readCount.ToString() + " Found:" + ItemsFound.ToString() + " not found:" + ItemsNotFound.ToString());

            btn.Enabled = true;
            Cursor = Cursors.Default;

        }

        Int32 readCount = 0;

        private void ValidateItems()
        {
            readCount = 0;
            ItemsFound = 0;
            ItemsNotFound = 0;

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
            string sBatchConNum = vRow["batchcontractforCIuk"].ToString();
            BatchConNum = 0;
            Int32.TryParse(sBatchConNum, out BatchConNum);
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
                ItemsFound += 1;
                lblFoundCount.Text = ItemsFound.ToString();
            }
            else
            {
                ItemsNotFound += 1;
                lblNotFoundCount.Text = ItemsNotFound.ToString();
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
                string tf = myRow["isDeFstinationFound"].ToString();
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

        private void CreateDeleteSQLForUnMatchedItems(string vFileName)
        {
            StreamWriter outputFile = new StreamWriter(vFileName);

            Int32 numFound = 0;
            Int32 numNotFound = 0;

            foreach (DataRow myRow in dtMismatches.Rows)
            {
                string tf = myRow["isDestinationFound"].ToString();
                string uk = myRow["contractitemsUK"].ToString();
                string badBatchNum = myRow["cibatchnum"].ToString();

                if (tf.ToLower() == "true")
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

            outputFile.WriteLine("--CreateDeleteSQLForUnMatchedItems Num Found:" + numFound.ToString() + " not found:" + numNotFound.ToString());
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
        #endregion End Mismatch


        #region "Missing Items"
        private void RetrieveDataForAddItems()
        {
            string fieldNameInvoiceNo = "auth_number";

            lblMissingStatus.Text = "Loading File";
            Application.DoEvents();

            DataTable dtTest = new DataTable();
            string[] cols = createAddItemColumns();
            SICommon.ColumnUtilities.addColumnsToTable(dtTest, cols);

            _SqlDataRoutines.ConnectionString = sConnectionString;

            Int64 readCount = 0;
            Int64 displayCount = 0;
            Int64 maxDisplayCount = 1;

            lblMissingStatus.Text = "Processing";
            Application.DoEvents();

            log.Info("Starting RetrieveDataForAddItems:" + DateTime.Now.ToString());

            ContractsWithItems = 0;
            ContractsWithNoItems = 0;
            ContractsFound = 0;
            ContractsNotFound = 0;
            ItemsFound = 0;
            ItemsNotFound = 0;

            string prevInvoiceNo = "";

            foreach (DataRow myRow in dtOutput.Rows)
            {
                bool isChangedInvoice = false;

                string InvoiceNo;
                InvoiceNo = myRow[fieldNameInvoiceNo].ToString();

                if (prevInvoiceNo != InvoiceNo)
                {
                    prevInvoiceNo = InvoiceNo;
                    seqnum = 1;
                    ContractsFound += 1;
                    isChangedInvoice = true;
                }

                displayCount += 1;
                readCount += 1;
                seqnum += 1;

                if (skipRecords > 0 && readCount < skipRecords)
                {
                    continue;
                }

                bool isItemFound = RetrieveContractInfoForAddItem(InvoiceNo, dtTest, myRow);

                if (isChangedInvoice)
                {
                    if (isItemFound)
                    {
                        ContractsWithItems += 1;
                    }
                    else
                    {
                        ContractsWithNoItems += 1;
                    }

                }

                if (displayCount >= maxDisplayCount)
                {
                    displayCount = 1;
                    lblMissingStatus.Text = readCount.ToString();
                    lblMissingStatus2.Text = "Contracts: " + ContractsFound.ToString() + "  Contracts No Items:" + ContractsWithNoItems.ToString() + " Items found:" + ItemsFound.ToString() + " No Items:" + ItemsNotFound.ToString();
                    Application.DoEvents();
                }
                if (readCount > maxToRead && maxToRead > 0)
                {
                    break;
                }
            }

            dgvProcessOutput.DataSource = dtTest;
            dgvProcessOutput.Refresh();
            
            log.Info("Finished RetrieveDataForAddItems:" + DateTime.Now.ToString());
            log.Info("Contracts found: " + ContractsFound.ToString() + " contracts not found: " + ContractsNotFound.ToString());
            log.Info("Contracts With Items: " + ContractsWithItems.ToString() + " Contract With NO Items: " + ContractsWithNoItems.ToString());
 
        }

        private bool RetrieveContractInfoForAddItem(string vInvoiceNo, DataTable vdt, DataRow vRow)
        {
            if (_SqlDataRoutines == null)
            {
                _SqlDataRoutines = new DbDataRoutines();
                _SqlDataRoutines.ConnectionString = sConnectionString;
            }

            string sCancelDate = "";
            string sDealer = "";
            string sStore = "";

            sCancelDate = vRow["Cancel_Date"].ToString();
            sDealer = vRow["Dealer"].ToString();
            sStore = vRow["Store_number"].ToString();

            string sqlFMT = "Select (select count(*) from contractitems where contractitems.batchconnum=batchcontract.uniquekey) as ItemCount,batchcontract.* From Batchcontract " +
                "inner join batchheader on batchheader.uniquekey = batchcontract.batchnum " +
                "inner join dealer on dealer.uniquekey = batchheader.dealer " +
                "where batchcontract.invoiceno = '{0}' " +
                "and dealer.dealernumber = '" + sDealer + "' " +
                //"and batchcontract.cancelentrydate is null " +
                "order by batchcontract.uniquekey desc";

            string sql = string.Format(sqlFMT, vInvoiceNo);

            _SqlDataRoutines.SQLString = sql;

            DataTable dt = _SqlDataRoutines.getTableFromDB();

            DataRow newRow = vdt.NewRow();

            if (vdt.Columns.Contains("invoiceNo"))
            {
                newRow["invoiceNo"] = vInvoiceNo;
            }
            else if (vdt.Columns.Contains("Auth_Number"))
            {
                newRow["Auth_Number"] = vInvoiceNo;
            }

            bool isItemFound = false;

            if (dt.Rows.Count == 0)
            {
                ContractsNotFound += 1;
                newRow["Certificate"] = "*";
                newRow["isValid"] = "N";

            }
            else
            {
                DataRow dr1 = dt.Rows[0];

                newRow["Certificate"] = dr1["Certificate"];
                newRow["BatchNumber"] = dr1["batchnum"].ToString();
                newRow["batchconnum"] = dr1["uniquekey"].ToString();
                newRow["Manufacturer"] = vRow["Manufacturer"].ToString();
                newRow["Model"] = vRow["Model"].ToString();
                newRow["Serial_Number"] = vRow["Serial_number"].ToString();
                newRow["Quantity"] = vRow["Quantity"].ToString();
                newRow["Purchase_Price"] = vRow["Purchase_Price"].ToString();
                newRow["Description"] = vRow["Description"].ToString();
                newRow["sItemID"] = vRow["ItemProductID"].ToString();

                newRow["FirstName"] = dr1["FirstName"];
                newRow["LastName"] = dr1["LastName"];
                newRow["sUnitran"] = dr1["sUniTran"];
                newRow["seqnum"] = seqnum.ToString();
                newRow["RecCount"] = dt.Rows.Count.ToString();
                string sItemCount = dr1["itemCount"].ToString();
                newRow["itemcount"] = sItemCount;

                if (sItemCount == "0")
                {
                    ItemsNotFound += 1;
                }
                else
                {
                    ItemsFound += 1;
                    isItemFound = true;
                }
                newRow["isValid"] = "Y";
            }

            vdt.Rows.Add(newRow);

            return isItemFound;
        }
        private string[] createAddItemColumns()
        {
            string[] columns =
            {
                "invoiceNo",
                "isValid",
                "itemCount",
                "batchconnum",
                "batchnumber",
                "Certificate",
                "manufacturer",
                "model",
                "description",
                "quantity",
                "serial_number",
                "purchase_price",
                "firstname",
                "lastName",
                "sItemID",
                "sUniTran",
                "seqnum",
                "reccount",
                "lastcolumn"

            };
            return columns;
        }
 
        private void btnLoadMissing_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.Save();
            Button btn = (Button)sender;
            btn.Enabled = false;
            Cursor = Cursors.WaitCursor;

            LoadProcessInputPremiumFile(sender, false);

            //RetrieveDataForAddItems(txtAddItemFileName.Text, txtAddItemSheetName.Text, true);
            if (isFatalError)
            {
                MessageBox.Show("Fatal Error Encounterd:" + FatalErrorMessage);
            }

            btn.Enabled = true;
            Cursor = Cursors.Default;
        }

        private void btnAddItemFindFile_Click(object sender, EventArgs e)
        {
            string sFileName = FindFileName(txtMissingFileName);
        }

        private void btnMissingRetrieveInfo_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.Save();
            Button btn = (Button)sender;
            btn.Enabled = false;
            Cursor = Cursors.WaitCursor;

            RetrieveDataForAddItems();
            if (isFatalError)
            {
                MessageBox.Show("Fatal Error Encounterd:" + FatalErrorMessage);
            }

            btn.Enabled = true;
            Cursor = Cursors.Default;
        }

        private ContractItemsInput BuildContractItems(DataRow vRow)
        {

            var myContractItems = new ContractItemsInput();

            Int32 iParm = 0;
            double iDouble = 0;

            Int32.TryParse(vRow["BatchNumber"].ToString(), out iParm);
            myContractItems.BatchNum = iParm;

            Int32.TryParse( vRow["BatchConNum"].ToString(),out iParm);
            myContractItems.BatchConNum = iParm;

            myContractItems.Description = vRow["Description"].ToString();

            double.TryParse(vRow["Purchase_Price"].ToString(), out iDouble);
            myContractItems.Cost = iDouble;
            myContractItems.MFG = vRow["Manufacturer"].ToString();
            myContractItems.Model = vRow["Model"].ToString();

            myContractItems.SerialNum = vRow["Serial_Number"].ToString();
            iParm = 0;
            Int32.TryParse(vRow["Quantity"].ToString(), out iParm);
            myContractItems.Quantity = iParm;
            iParm = 0;
            Int32.TryParse(vRow["seqnum"].ToString(), out iParm);
            myContractItems.SeqNum = iParm;

            //myContractItems.LastUpdate = LastUpdate;

            myContractItems.sItemID = vRow["sItemID"].ToString();
            myContractItems.sUniTran = vRow["sUniTran"].ToString();
            //myContractItems.lastUser = lastUser;

            return myContractItems;
        }

        private void btnMissingWriteContractItems_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.Save();
            if (txtMissingOutputFile.Text.Trim().Length == 0)
            {
                MessageBox.Show("Please enter an output file name before proceeding - Retry");
                return;
            }

            Button btn = (Button)sender;
            btn.Enabled = false;
            Cursor = Cursors.WaitCursor;
            
            WriteContractItems(txtMissingOutputFile.Text.Trim());

            btn.Enabled = true;
            Cursor = Cursors.Default;

            MessageBox.Show("Output file written!");

        }


        private void WriteContractItems(string vFileName)
        {
            StreamWriter outputFile = new StreamWriter(vFileName);

            DataTable dt = (DataTable)dgvProcessOutput.DataSource;

            foreach (DataRow myRow in dt.Rows)
            {
                string sRecCount = myRow["itemCount"].ToString();
                if (sRecCount == "")
                {
                    continue;
                }

                Int32 recCount = 0;
                Int32.TryParse(sRecCount, out recCount);

                if (recCount > 0)
                {
                    continue;
                }

                ContractItemsInput myCI = BuildContractItems(myRow);

                myCI.LastUpdate = currentUpdateTime;

                string sql = BuildContractItemInsertSQL(myCI);

                //outputFile.WriteLine("output:" + myCI.SeqNum.ToString() + ":" + myCI.BatchConNum + " desc:" + myCI.Description + " " + myCI.Cost.ToString());
                outputFile.WriteLine(sql);
            }

            outputFile.Close();
        }
        public string BuildContractItemInsertSQL(ContractItemsInput vContractItem)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("INSERT INTO contractitems ");
            sb.Append("(");
            sb.Append("BatchNum");
            sb.Append(", ");
            sb.Append("BatchConNum");
            sb.Append(", ");
            sb.Append("Description");
            sb.Append(", ");
            sb.Append("Cost");
            sb.Append(", ");
            sb.Append("MFG");
            sb.Append(", ");
            sb.Append("Model");
            sb.Append(", ");
            sb.Append("SerialNum");
            sb.Append(", ");
            sb.Append("Quantity");
            sb.Append(", ");
            sb.Append("SeqNum");
            sb.Append(", ");
            sb.Append("LastUpdate");
            sb.Append(", ");
            sb.Append("sItemID");
            sb.Append(", ");
            sb.Append("sUniTran");
            sb.Append(", ");
            sb.Append("lastUser");

            sb.Append(") ");
            sb.Append("VALUES (");
            sb.Append(vContractItem.BatchNum);
            sb.Append(", ");
            sb.Append(vContractItem.BatchConNum);
            sb.Append(", ");
            sb.Append("'");
            sb.Append(SQLUtil.Encode(vContractItem.Description));
            sb.Append("'");
            sb.Append(", ");
            sb.Append(vContractItem.Cost);
            sb.Append(", ");
            sb.Append("'");
            sb.Append(SQLUtil.Encode(vContractItem.MFG));
            sb.Append("'");
            sb.Append(", ");
            sb.Append("'");
            sb.Append(SQLUtil.Encode(vContractItem.Model));
            sb.Append("'");
            sb.Append(", ");
            sb.Append("'");
            sb.Append(SQLUtil.Encode( vContractItem.SerialNum));
            sb.Append("'");
            sb.Append(", ");
            sb.Append(vContractItem.Quantity);
            sb.Append(", ");
            sb.Append(vContractItem.SeqNum);
            sb.Append(", ");
            sb.Append("'");
            sb.Append(SICommon.SQLUtil.ToDBDateTime( vContractItem.LastUpdate));
            sb.Append("'");
            sb.Append(", ");
            sb.Append("'");
            sb.Append(SICommon.SQLUtil.Encode( vContractItem.sItemID));
            sb.Append("'");
            sb.Append(", ");
            sb.Append("'");
            sb.Append(vContractItem.sUniTran);
            sb.Append("'");
            sb.Append(", ");
            sb.Append("'");
            sb.Append("fixcontractItem");
            sb.Append("'");

            sb.Append(")");

            return sb.ToString();
        }

        #endregion Missing items

        private void btnUnmatchedDelete_Click(object sender, EventArgs e)
        {
            string fileName = txtOutputFileName.Text.Trim();

            if (fileName.Length == 0)
            {
                MessageBox.Show("You must enter an output file name - Retry");
                return;
            }

            Properties.Settings.Default.Save();

            CreateDeleteSQLForUnMatchedItems(fileName);

            MessageBox.Show("Output file written");

        }
    }
}