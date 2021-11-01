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
using ContractExcelToXml;
using System.Diagnostics;

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
        string Version = "1.0";
        DataTable dtMismatches;
        DataTable dtClaimCarrierMismatches;
        LineItems _lineItem;
        private string FatalErrorMessage;
        private bool isFatalError;
        private DbDataRoutines _SqlDataRoutines;
        DataTable dtInput;
        DataTable dtOutput;
        DataTable dtContractItemsUpdateInput;
        DataTable dtContractItemsUpdateOutput;
        Int32 ContractsWithNoItems = 0;
        Int32 ContractsWithItems = 0;
        Int32 ContractsNotFound = 0;
        Int32 ContractsFound = 0;
        Int32 ItemsFound = 0;
        Int32 ItemsNotFound = 0;
        Int32 ItemDetailsFound = 0;
        Int32 ItemDetailsNotFound = 0;
        Int32 ItemsDetailsMatch = 0;
        Int32 ItemsDetailsNotMatch = 0;
        Int32 seqnum = 0;
        DateTime currentUpdateTime = DateTime.Now;
        Int32 skipRecords = 00;
        string UpdateItemsOption = "UpdateItems";
        Int32 ItemsNeedUpdate = 0;
        Int32 ItemsNotNeedUpdate = 0;

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
            log.Debug("ConnectionString:" + SQLUtil.HideField(sConnectionString, "pwd"));

            _lineItem = new LineItems();
            _lineItem.ConnectionStringSQL = sConnectionString;

            _SqlDataRoutines = new DbDataRoutines(sConnectionString);
            _SqlDataRoutines.DefaultCommandTimeout = 600;

            lblVersion.Text = Version;

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
        private bool WriteExcelFile(TextBox vFileName, TextBox vSheetName, DataTable vdt, bool vIsWriteHeaders, ProcessExcel vProcessExcel = null)
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
                            oAdapter.SelectCommand.CommandTimeout = 600;
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
            string option = "Missing";
            if (btn.Name == "btnContractItemsUpdateLoad")
            {
                option = "Update";
            }
            string InputExcelFileName;
            string InputSheetName;

            if (option == "Update")
            {
                InputExcelFileName = txtContractItemsUpdateFileName.Text.Trim();
                InputSheetName = txtContractItemsUpdateSheetName.Text.Trim();
                lblContractItemsUpdateStatus.Text = "Loading " + InputExcelFileName;
            }
            else
            {
                InputExcelFileName = txtMissingFileName.Text.Trim();
                InputSheetName = txtMissingSheetName.Text.Trim();
                lblMissingStatus.Text = "Loading " + InputExcelFileName;

            }

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

            if (option == "Update")
            {
                dtContractItemsUpdateInput = _process.inputDataTable;
                dtContractItemsUpdateOutput = _process.outputDataTable;

                dgvContractItemsUpdateInput.DataSource = _process.outputDataTable;
                dgvContractItemsUpdateInput.Refresh();

                dgvContractItemsUpdateOutput.DataSource = _process.outputDataTable;
                dgvContractItemsUpdateOutput.Refresh();

                lblContractItemsUpdateOutputRecordCount.Text = _process.outputDataTable.Rows.Count.ToString();
                lblContractItemsUpdateStatus.Text = "Finished";
            }
            else
            {
                //dtLookupProcess = _process.outputDataTable;
                dtInput = _process.inputDataTable;
                dtOutput = _process.outputDataTable;

                dgvProcessInput.DataSource = _process.outputDataTable;
                dgvProcessInput.Refresh();

                dgvProcessOutput.DataSource = _process.outputDataTable;
                dgvProcessOutput.Refresh();

                lblProcessOutputCount.Text = _process.outputDataTable.Rows.Count.ToString();
                lblMissingStatus.Text = "Finished";
            }

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

            ValidateItems(dtMismatches);

            log.Info("Items: " + readCount.ToString() + " Found:" + ItemsFound.ToString() + " not found:" + ItemsNotFound.ToString());

            btn.Enabled = true;
            Cursor = Cursors.Default;

        }

        Int32 readCount = 0;

        private void ValidateItems(DataTable vdt)
        {
            readCount = 0;
            ItemsFound = 0;
            ItemsNotFound = 0;

            foreach (DataRow myRow in vdt.Rows)
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
            _contractItem.BatchConNum = (Int32)vRow["batchConNum"];
            _contractItem.Model = vRow["Model"].ToString();
            //_contractItem.Model = "";    //blank model if it is different but not significant in finding the right contract
            _contractItem.SerialNum = vRow["SerialNum"].ToString();
            _contractItem.SerialNum = vRow["sItemID"].ToString();
            _contractItem.SerialNum = vRow["sUniTran"].ToString();
            _contractItem.Quantity = (Int32)vRow["Quantity"];
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
        private void RetrieveDataForAddItems(DataTable vdt, string vOption = "")
        {
            string fieldNameInvoiceNo = "auth_number";

            Label lStatus = lblMissingStatus;
            Label lStatus2 = lblMissingStatus2;
            DataGridView dgv = dgvProcessOutput;

            if (vOption == UpdateItemsOption)
            {
                lStatus = lblContractItemsUpdateStatus;
                lStatus2 = lblContractItemsUpdateStatus2;
                dgv = dgvContractItemsUpdateOutput;
            }

            lStatus.Text = "Loading File";

            Application.DoEvents();

            DataTable dtTest = new DataTable();
            string[] cols = createAddItemColumns(vOption);
            SICommon.ColumnUtilities.addColumnsToTable(dtTest, cols);

            _SqlDataRoutines.ConnectionString = sConnectionString;

            Int64 readCount = 0;
            Int64 displayCount = 0;
            Int64 maxDisplayCount = 1;

            lStatus.Text = "Processing";
            Application.DoEvents();

            log.Info("Starting RetrieveDataForAddItems(" + vOption + ") :" + DateTime.Now.ToString());

            ContractsWithItems = 0;
            ContractsWithNoItems = 0;
            ContractsFound = 0;
            ContractsNotFound = 0;
            ItemsFound = 0;
            ItemsNotFound = 0;
            ItemsDetailsMatch = 0;
            ItemsDetailsNotMatch = 0;
            ItemsNeedUpdate = 0;
            ItemsNotNeedUpdate = 0;

            string prevInvoiceNo = "";

            foreach (DataRow myRow in vdt.Rows)
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

                if (vOption == UpdateItemsOption)
                {
                    Int32 i = dtTest.Rows.Count;
                    DataRow newRow = dtTest.Rows[i - 1];

                    RetrieveContractItem(newRow);

                    if (newRow["itemNeedsUpdate"].ToString() == "Y")
                    {
                        ItemsNeedUpdate += 1;
                    }
                    else
                    {
                        ItemsNotNeedUpdate += 1;
                    }
                }
                    if (displayCount >= maxDisplayCount)
                    {
                        displayCount = 1;
                        lStatus.Text = readCount.ToString();
                        if (vOption == UpdateItemsOption)
                        {
                            lStatus2.Text = "Contracts: " + ContractsFound.ToString()
                                + "  Contracts No Items:" + ContractsWithNoItems.ToString()
                                + " Items found:" + ItemsFound.ToString()
                                + " No Items:" + ItemsNotFound.ToString()
                                + " Details Match:" + ItemsDetailsMatch.ToString()
                                + " No Match:" + ItemsDetailsNotMatch.ToString()
                                + " Needs Update:" + ItemsNeedUpdate.ToString()
                                + " No Update Needed:" + ItemsNotNeedUpdate.ToString();

                        }
                        else
                        {
                            lStatus2.Text = "Contracts: " + ContractsFound.ToString()
                                + "  Contracts No Items:" + ContractsWithNoItems.ToString()
                                + " Items found:" + ItemsFound.ToString()
                                + " No Items:" + ItemsNotFound.ToString();

                        }
                        Application.DoEvents();
                    }

                    if (readCount > maxToRead && maxToRead > 0)
                    {
                        break;
                    }
                }

                dgv.DataSource = dtTest;
                dgv.Refresh();

                log.Info("Finished RetrieveDataForAddItems(" + vOption + "):" + DateTime.Now.ToString());
                log.Info("Contracts found: " + ContractsFound.ToString() + " contracts not found: " + ContractsNotFound.ToString());
                log.Info("Contracts With Items: " + ContractsWithItems.ToString() + " Contract With NO Items: " + ContractsWithNoItems.ToString());


        }

        private bool RetrieveContractItem(DataRow vRow)
        {
            bool isFound = false;
            if (_SqlDataRoutines == null)
            {
                _SqlDataRoutines = new DbDataRoutines();
                _SqlDataRoutines.ConnectionString = sConnectionString;
            }

            ItemDetailsFound = 0;
            ItemDetailsNotFound = 0;
            string itemTotalCount = vRow["itemCount"].ToString();

            string inItemId;
            inItemId = vRow["sItemId"].ToString();

            string inDesc = vRow["Description"].ToString();

            string inMfg = vRow["Manufacturer"].ToString();
            string inModel = vRow["Model"].ToString();
            string inSerialnum = vRow["Serial_Number"].ToString();
            string sBatchconnum = vRow["batchconnum"].ToString();
            string inCost = vRow["Purchase_price"].ToString();
            string inInvoiceNo = vRow["InvoiceNo"].ToString();

            if (sBatchconnum.Trim().Length == 0)
            {
                vRow["isItemFound"] = "N";
                isFound = false;
                return isFound;
            }

            string sqlFMT;
            if (itemTotalCount == "1")
            {
                sqlFMT = $"Select * from contractitems where contractitems.batchconnum = {sBatchconnum} ";
            }
            else
            {
                sqlFMT = $"Select * from contractitems where contractItems.batchconnum = {sBatchconnum} " +
                    // $"and mfg = '{inMfg}' and model = '{inModel}' and serialNum = '{inSerialnum}' and sitemid <> '{inItemId}'";
                    $"and mfg = '{inMfg}' and model = '{inModel}' and serialNum = '{inSerialnum}' and cost = '{inCost}' and description = '{inDesc}'";
            }


            string sql = sqlFMT;

            _SqlDataRoutines.SQLString = sql;

            DataTable dt = _SqlDataRoutines.getTableFromDB();

            if (dt.Rows.Count == 0)
            {
                vRow["isItemFound"] = "N";
                isFound = false;
                ItemDetailsNotFound += 1;
                vRow["detailsMatch"] = "N";
                ItemsDetailsNotMatch += 1;
                vRow["itemNeedsUpdate"] = "*";

            }
            else
            {
                vRow["isItemFound"] = dt.Rows.Count.ToString();
                isFound = true;
                ItemDetailsFound += 1;

            }

            if (!isFound)
            {
                return isFound;
            }

            Int32 itemsMatch = 0;

            foreach (DataRow myRow in dt.Rows)
            {

                string ciUK = myRow["uniquekey"].ToString();

                string ciItemID = myRow["sItemID"].ToString();
                string desc = myRow["Description"].ToString();
                string mfg = myRow["MFG"].ToString();
                string model = myRow["Model"].ToString();
                string serialnum = myRow["SerialNum"].ToString();
                string lastupdate = myRow["Lastupdate"].ToString();
                string cost = myRow["cost"].ToString();

                vRow["ciMFG"] = mfg;
                vRow["ciModel"] = model;
                vRow["ciDesc"] = desc;
                vRow["ciQTY"] = myRow["quantity"].ToString();
                vRow["ciItemID"] = ciItemID;
                vRow["ciSerial"] = serialnum;
                vRow["ciCost"] = myRow["Cost"].ToString();
                vRow["ciLastUpdate"] = lastupdate;

                string sItemNeedsUpdate = "Y";

                if (inItemId == ciItemID)
                {
                    vRow["itemNeedsUpdate"] = "N";
                    sItemNeedsUpdate = "N";
                }
                else
                {
                    vRow["itemNeedsUpdate"] = "Y";
                }

                vRow["ciUK"] = ciUK;

                double dCost = 0;
                double.TryParse(cost, out dCost);
                double dInCost = 0;
                double.TryParse(inCost, out dInCost);

                if (desc == inDesc
                    && mfg == inMfg
                    && model == inModel
                    && serialnum == inSerialnum
                    && dCost == dInCost
                    )

                {
                    itemsMatch += 1;
                    if (sItemNeedsUpdate == "N")
                    {
                        break;
                    }

                }
                else
                {
                    string x = "dummy";
                }
            }

            if (itemsMatch == 0)
            {
                vRow["detailsMatch"] = "N";
                ItemsDetailsNotMatch += 1;
                vRow["itemNeedsUpdate"] = "*";

            }
            else
            {
                vRow["detailsMatch"] = itemsMatch.ToString();
                ItemsDetailsMatch += 1;
            }

            return isFound;
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

            if (sDealer.Length == 0)
            {
                sDealer = "25398";
            }
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
                newRow["entrydate"] = dr1["entrydate"].ToString();

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
        private string[] createAddItemColumns(string vOption="")
        {
            string[] columns; 
            if (vOption == UpdateItemsOption)
            {
                columns = new string []
                    {
                    "invoiceNo",
                    "isValid",
                    "itemCount",
                    "isItemFound",
                    "itemNeedsUpdate",
                    "batchconnum",
                    "detailsMatch",
                    "entrydate",
                    "sItemID",
                    "ciItemId",
                    "ciCost",
                    "purchase_price",
                    "ciUK",
                    "ciDesc",
                    "ciModel",
                    "ciMFG",
                    "ciSerial",
                    "ciQty",
                    "ciLastUpdate",
                    "batchnumber",
                    "Certificate",
                    "manufacturer",
                    "model",
                    "description",
                    "quantity",
                    "serial_number", 
                    "firstname",
                    "lastName",
 
                    "sUniTran",
                    "seqnum",
                    "reccount",
                    "lastcolumn"
                };
            }
            else
            {
                columns = new string []
                {
                    "invoiceNo",
                    "isValid",
                    "itemCount",
                    "batchnumber",
                    "batchconnum",
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
            }

            return columns;
        }
 
        private void btnLoadMissing_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.Save();
            Button btn = (Button)sender;
            btn.Enabled = false;
            Cursor = Cursors.WaitCursor;

            LoadProcessInputPremiumFile(sender, false);

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

            RetrieveDataForAddItems(dtOutput, "Missing");
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

        private void btnRetrieveMissingItems_Click(object sender, EventArgs e)
        {
            Button btn = (Button)sender;
            btn.Enabled = false;
            Cursor = Cursors.WaitCursor;

            dtMismatches = RetrieveMissingContractItems();

            lblMismatchedRecordCount.Text = dtMismatches.Rows.Count.ToString();

            dgvProcessOutput.DataSource = dtMismatches;
            dgvProcessOutput.Refresh();

            btn.Enabled = true;
            Cursor = Cursors.Default;
        }

        private DataTable RetrieveMissingContractItems()
        {
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT ");
            sbSQL.Append("count(*) as count, ");
            sbSQL.Append("convert(date,batchcontract.EntryDate) as Entrydate, ");
            sbSQL.Append("dealergroup.sGroupNumber, ");
            sbSQL.Append("min(batchcontract.BatchNum) as minBatchNum, ");
            sbSQL.Append("max(batchcontract.BatchNum) as maxBatchNum, ");
            sbSQL.Append("min(batchcontract.uniquekey) as minBCUK, ");
            sbSQL.Append("max(batchcontract.uniquekey) as maxBCUK, ");
            sbSQL.Append("min(batchcontract.Certificate) as minCertNo, ");
            sbSQL.Append("max(batchcontract.Certificate) as maxCertNo, ");
            sbSQL.Append("min(batchcontract.ExpirationDate) as minExpiration, ");
            sbSQL.Append("max(batchcontract.ExpirationDate) AS maxExpiration, ");
            sbSQL.Append("min(dealer.dealernumber) as minDealerNumber, ");
            sbSQL.Append("max(dealer.DealerNumber) as maxDealerNumber, ");
            sbSQL.Append("min(batchcontract.entrydate) as minEntrydate, ");
            sbSQL.Append("max(batchcontract.entrydate) as maxEntrydate ");
            sbSQL.Append("FROM batchcontract ");
            sbSQL.Append("left JOIN contractitems ON batchcontract.uniquekey=contractitems.BatchConNum ");
            sbSQL.Append("INNER JOIN batchheader bh ON bh.uniquekey=batchcontract.BatchNum ");
            sbSQL.Append("INNER JOIN dealer ON dealer.uniquekey=bh.Dealer ");
            sbSQL.Append("INNER JOIN dealergroupSELECT ON dealergroupselect.iDealer = dealer.uniquekey ");
            sbSQL.Append("INNER JOIN dealergroup ON dealergroup.uniquekey = dealergroupselect.iDealerGroup ");
            sbSQL.Append("WHERE  contractitems.uniquekey is null ");
            sbSQL.Append("and convert(date,entrydate) >= '2021-01-01' ");
            sbSQL.Append("and convert(date,expirationdate) >= '2021-01-01' ");
            sbSQL.Append("and cancelentrydate is null ");
            sbSQL.Append("and sgroupnumber <> 'z4wcportal' ");
            sbSQL.Append("GROUP BY convert(date,batchcontract.EntryDate), dealergroup.sGroupNumber ");
            sbSQL.Append("ORDER BY convert(date,batchcontract.EntryDate) desc ");
            sbSQL.Append(" ");
            sbSQL.Append(" ");

            DataTable dt;
            _SqlDataRoutines.SQLString = sbSQL.ToString();
            dt = _SqlDataRoutines.getTableFromDB();
            return dt;
        }

        private DataTable RetrieveClaimCarrierMismatches()
        {
            StringBuilder sbSQL = new StringBuilder();

            sbSQL = new StringBuilder();
            sbSQL.Append("SELECT ");
            sbSQL.Append("claimmaster.iWarranty, ");
            sbSQL.Append("claimmaster.Certificate as ClaimCertificate, ");
            sbSQL.Append("claimmaster.uniquekey as claimnumber, ");
            sbSQL.Append("claimmaster.createdate as claimCreateDate, ");
            sbSQL.Append("batchcontract.icarrier as bccarrier, ");
            sbSQL.Append("InsuranceKey as claimcarrier, ");
            sbSQL.Append("claimstatus, ");
            sbSQL.Append("closedate, ");
            sbSQL.Append("batchcontract.entrydate as bcEntryDate, ");
            sbSQL.Append("(SELECT count(*) FROM payments WHERE payments.iClaim = claimmaster.uniquekey) as payments, ");
            sbSQL.Append("(SELECT sum(payments.checkamount) FROM payments WHERE payments.iClaim = claimmaster.uniquekey) as paymentAmt, ");
            sbSQL.Append("batchcontract.LastName as bcLast, ");
            sbSQL.Append("claimmaster.LName as claimLast, ");
            sbSQL.Append("claimmaster.lastuser as claimLastUer, ");
            sbSQL.Append("claimmaster.LossDate, ");
            sbSQL.Append("claimmaster.InvoiceNum as claimInvoicenumber, ");
            sbSQL.Append("batchcontract.InvoiceNo as bcInvoiceNo, ");
            sbSQL.Append("batchcontract.Certificate, ");
            sbSQL.Append("carrier.carriername as bccarrierName, ");
            sbSQL.Append("carrier2.CarrierName as claimCarrierName ");
            sbSQL.Append("FROM ");
            sbSQL.Append("claimmaster ");
            sbSQL.Append("LEFT JOIN batchcontract ON batchcontract.uniquekey = claimmaster.iWarranty ");
            sbSQL.Append("left JOIN carrier ON carrier.uniquekey = batchcontract.icarrier ");
            sbSQL.Append("left JOIN carrier carrier2 ON carrier2.uniquekey = claimmaster.InsuranceKey ");
            sbSQL.Append("WHERE ");
            sbSQL.Append("insurancekey <> icarrier ");
            sbSQL.Append("and convert(date,createdate) > '2020-01-01' ");
            sbSQL.Append("ORDER BY claimmaster.createdate desc ");
            sbSQL.Append(" ");

            DataTable dt;
            _SqlDataRoutines.SQLString = sbSQL.ToString();
            dt = _SqlDataRoutines.getTableFromDB();
            return dt;

        }
        private void btnClaimCarrierFixRetrieve_Click(object sender, EventArgs e)
        {
            Button btn = (Button)sender;
            btn.Enabled = false;
            Cursor = Cursors.WaitCursor;

            dtClaimCarrierMismatches = RetrieveClaimCarrierMismatches();

            lblClaimCarrierMismatchOutputCount.Text = dtClaimCarrierMismatches.Rows.Count.ToString();

            dgvClaimCarrierMismatch.DataSource = dtClaimCarrierMismatches;
            dgvClaimCarrierMismatch.Refresh();

            btn.Enabled = true;
            Cursor = Cursors.Default;
        }

        private void btnWriteClaimCarrierMismatchOutput_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            Button btn = (Button)sender;
            btn.Enabled = false;

            bool isWriteHeaders = true;

            WriteExcelFile(txtClaimCarrierFixOutputName, txtClaimCarrierMismatchSheetName, dtClaimCarrierMismatches, isWriteHeaders);

            Cursor = Cursors.Default;
            btn.Enabled = true;
        }

        private void btnClaimCarrierMismatchExpand_Click(object sender, EventArgs e)
        {
            ExpandDataGridView((Button)sender, dgvClaimCarrierMismatch);

        }

        private void btnClaimCarrierFixWriteSQL_Click(object sender, EventArgs e)
        {
            string fileName = txtClaimCarrierFixOutputFileName.Text.Trim();

            if (fileName.Length == 0)
            {
                MessageBox.Show("You must enter an output file name - Retry");
                return;
            }

            Properties.Settings.Default.Save();

            CreateSQLForClaimCarrierFix(fileName);

            MessageBox.Show("Output file written");
        }

        private void CreateSQLForClaimCarrierFix(string vFileName)
        {
            StreamWriter outputFile = new StreamWriter(vFileName);

            Int32 numFound = 0;
            Int32 numNotFound = 0;

            foreach (DataRow myRow in dtClaimCarrierMismatches.Rows)
            {
                string uk = myRow["iWarranty"].ToString();
                string bcCarrier = myRow["bccarrier"].ToString();
                string claimcarrier = myRow["claimcarrier"].ToString();
                string claimNumber = myRow["claimnumber"].ToString();

                numFound += 1;

                string sSQLFmt = $"update claimmaster set insurancekey = '{bcCarrier}' where uniquekey = {claimNumber} and insurancekey='{claimcarrier}';";
                outputFile.WriteLine(sSQLFmt);

            }

            outputFile.WriteLine("--Num Found:" + numFound.ToString() + " not found:" + numNotFound.ToString());
            outputFile.Close();

        }



        private void btnContractItemsUpdateLoad_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.Save();
            Button btn = (Button)sender;
            btn.Enabled = false;
            Cursor = Cursors.WaitCursor;

            LoadProcessInputPremiumFile(sender, false);

            if (isFatalError)
            {
                MessageBox.Show("Fatal Error Encounterd:" + FatalErrorMessage);
            }

            btn.Enabled = true;
            Cursor = Cursors.Default;
        }

        private void btnContractItemsUpdateRetrieve_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.Save();
            Button btn = (Button)sender;
            btn.Enabled = false;
            Cursor = Cursors.WaitCursor;

            System.Diagnostics.Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();

            RetrieveDataForAddItems(dtContractItemsUpdateOutput, "UpdateItems");

            stopWatch.Stop();
            TimeSpan elapsed = stopWatch.Elapsed;

            //ValidateItems(dtContractItemsUpdateOutput);
            lblContractItemsUpdateStatus.Text = "Elapsed:" + elapsed.ToString();

            if (isFatalError)
            {
                MessageBox.Show("Fatal Error Encounterd:" + FatalErrorMessage);
            }

            btn.Enabled = true;
            Cursor = Cursors.Default;
        }

        private void btnContractItemsUpdateWrite_Click(object sender, EventArgs e)
        {
            WriteContractItemsUpdate(txtContractItemsUpdateWriteSQL.Text.Trim());

            MessageBox.Show("Finished Writing File");
        }

        private void btnContractItemsUpdateOutputWrite_Click(object sender, EventArgs e)
        {

        }

        private void btnContractItemsUpdateOutputExpand_Click(object sender, EventArgs e)
        {
            ExpandDataGridView((Button)sender, dgvContractItemsUpdateOutput);
        }

        private void btnContractItemsUpdateInputWrite_Click(object sender, EventArgs e)
        {

        }

        private void btnContractItemsUpdateInputExpand_Click(object sender, EventArgs e)
        {

        }

        private void btnContractItemsUpdateFindFile2_Click(object sender, EventArgs e)
        {
                string sFileName = FindFileName(txtContractItemsUpdateFileName);
        }

        private string WriteContractItemsUpdate(string vFileName, string vOption="")
        {
            string sSQL = "";

            StreamWriter outputFile = new StreamWriter(vFileName);
            DataTable dt = (DataTable) dgvContractItemsUpdateOutput.DataSource;
            Int32 numFound = 0;
            Int32 numNotFound = 0;

            foreach (DataRow myRow in dt.Rows)
            {
                
                string uk = myRow["ciUK"].ToString();
                string sItemID = myRow["sItemID"].ToString();
                string isDetailsMatch = "";
                isDetailsMatch = myRow["detailsMatch"].ToString();
                if (isDetailsMatch == "N")
                {
                    numNotFound += 1;
                    continue;
                }
                string itemNeedUpdate = myRow["ItemNeedsUpdate"].ToString();
                numFound += 1;

                if (itemNeedUpdate != "Y")
                {
                    continue;
                }

                string desc = myRow["Description"].ToString();
                string batchconnum = myRow["batchconnum"].ToString();
                string mfg = myRow["Manufacturer"].ToString();
                string model = myRow["Model"].ToString();
                string serialNum = myRow["Serial_Number"].ToString();
                string inCost = myRow["purchase_price"].ToString();
                string sSQLFmt;
                if (vOption == "")
                {
                    sSQLFmt = $"update  contractitems set sItemid = '{sItemID}' " +
                        $" where batchconnum = '{batchconnum}' and description = '{desc}' and mfg = '{mfg}' and model = '{model}' and serialnum = '{serialNum}' and cost = '{inCost}';";
                }
                else
                {
                    sSQLFmt = $"update top(1) contractitems set sItemid = '{sItemID}' " +
                        $" where batchconnum = '{batchconnum}' and description = '{desc}' and mfg = '{mfg}' and model = '{model}' and serialnum = '{serialNum}' and cost = '{inCost}';";

                }
                outputFile.WriteLine(sSQLFmt);

            }

            outputFile.WriteLine("--Num Found:" + numFound.ToString() + " not found:" + numNotFound.ToString());
            outputFile.Close();

            return sSQL;
        }

        private void btnContractItemsUpdateErrorWrite_Click(object sender, EventArgs e)
        {
            WriteContractItemsUpdateErrors(txtContractItemsUpdateWriteSQL.Text);

            MessageBox.Show("Finished Writing File");

        }
        private string WriteContractItemsUpdateErrors(string vFileName)
        {
            string sSQL = "";

            StreamWriter outputFile;
            //outputFile = new StreamWriter(vFileName);
            DataTable dt = (DataTable)dgvContractItemsUpdateOutput.DataSource;
            Int32 numFound = 0;
            Int32 numNotFound = 0;

            BuildCSV _build = new BuildCSV();
 
            _build.CreateCSVHeaders(dt);
            _build.OpenCSVFile(vFileName);

            _build.WriteCSVLine(_build.CreateCSVHeaders(dt));

            outputFile = _build.CSVFile;

            foreach (DataRow myRow in dt.Rows)
            {

                string uk = myRow["ciUK"].ToString();
                string sItemID = myRow["sItemID"].ToString();
                string isDetailsMatch = "";
                isDetailsMatch = myRow["detailsMatch"].ToString();
                if (!(isDetailsMatch == "N"  || isDetailsMatch == "*" ))
                {
                   numFound += 1;
                   continue;
                }

                numNotFound += 1;

                string sSQLFmt;

                sSQLFmt =(string) _build.CreateCSVLine(myRow);

                _build.WriteCSVLine(sSQLFmt);

            }

            outputFile.WriteLine("--Num Found:" + numFound.ToString() + " not found :" + numNotFound.ToString());
            outputFile.Close();

            return sSQL;
        }

        private void btnContractItemsUpdateWrite2_Click(object sender, EventArgs e)
        {
            WriteContractItemsUpdate(txtContractItemsUpdateWriteSQL.Text.Trim(),"2");

            MessageBox.Show("Finished Writing File");
        }
    }
}
