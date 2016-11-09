using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Net;
using System.Text;
using System.Windows.Forms;
using System.IO;
using AdvSeleniumAPI;
using iATester;
using Model;
using Service;

public partial class Form1 : Form, iATester.iCom
{
    private delegate void DataGridViewCtrlAddDataRow(DataGridViewRow i_Row);
    private DataGridViewCtrlAddDataRow m_DataGridViewCtrlAddDataRow;
    internal const int Max_Rows_Val = 65535;

    IHttpReqService HttpReqService;
    DataHandleService dataHld;
    DeviceModel dev;
    string AddressIP = "", devName = "", path = "", browser = ""; bool ConnectFlg = false;
    int errorCnt = 0, step = 0;
    AdvSeleniumAPIv2 selenium; IniDataFmt LoadPara;
    //iATester
    //Send Log data to iAtester
    public event EventHandler<LogEventArgs> eLog = delegate { };
    //Send test result to iAtester
    public event EventHandler<ResultEventArgs> eResult = delegate { };
    //Send execution status to iAtester
    public event EventHandler<StatusEventArgs> eStatus = delegate { };
    //
    public Form1()
    {
        InitializeComponent();
    }

    private void Form1_Load(object sender, EventArgs e)
    {
        #region -- dataview --
        dataGridView1.ColumnHeadersVisible = true;
        DataGridViewTextBoxColumn newCol = new DataGridViewTextBoxColumn(); // add a column to the grid
        newCol.HeaderText = "Time Stamp";
        newCol.Name = "clmTs";
        newCol.Visible = true;
        newCol.Width = 90;
        dataGridView1.Columns.Add(newCol);
        newCol = new DataGridViewTextBoxColumn();
        newCol.HeaderText = "Exe Step";
        newCol.Name = "clmStp";
        newCol.Visible = true;
        newCol.Width = 150;
        dataGridView1.Columns.Add(newCol);
        newCol = new DataGridViewTextBoxColumn();
        newCol.HeaderText = "Result";
        newCol.Name = "clmRes";
        newCol.Visible = true;
        newCol.Width = 80;
        dataGridView1.Columns.Add(newCol);
        newCol = new DataGridViewTextBoxColumn();
        newCol.HeaderText = "Exe Time (ms)";
        newCol.Name = "clmExt";
        newCol.Visible = true;
        newCol.Width = 100;
        dataGridView1.Columns.Add(newCol);
        newCol = new DataGridViewTextBoxColumn();
        newCol.HeaderText = "Error Code";
        newCol.Name = "clmErr";
        newCol.Visible = true;
        newCol.Width = 200;
        dataGridView1.Columns.Add(newCol);

        for (int i = 0; i < dataGridView1.Columns.Count - 1; i++)
        {
            dataGridView1.Columns[i].SortMode = DataGridViewColumnSortMode.Automatic;
        }
        dataGridView1.Rows.Clear();
        try
        {
            m_DataGridViewCtrlAddDataRow = new DataGridViewCtrlAddDataRow(DataGridViewCtrlAddNewRow);
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.ToString());
        }
        #endregion

        HttpReqService = new HttpReqService();
        dataHld = new DataHandleService();
        LoadFile();
    }
    private void LoadFile()
    {
        LoadPara = dataHld.GetPara(Application.StartupPath);
        textBox1.Text = AddressIP = LoadPara.IP;
    }
    public void StartTest()//iATester
    {
        LoadFile();
        if (ExeConnectionDUT())
        {
            eStatus(this, new StatusEventArgs(iStatus.Running));
            WorkSteps();
            if (errorCnt > 0 || step != 99)
                eResult(this, new ResultEventArgs(iResult.Fail));
            else
                eResult(this, new ResultEventArgs(iResult.Pass));
        }
        else
            eResult(this, new ResultEventArgs(iResult.Fail));
        //
        eStatus(this, new StatusEventArgs(iStatus.Completion));
        Application.DoEvents();
    }


    private void DataGridViewCtrlAddNewRow(DataGridViewRow i_Row)
    {
        if (this.dataGridView1.InvokeRequired)
        {
            this.dataGridView1.Invoke(new DataGridViewCtrlAddDataRow(DataGridViewCtrlAddNewRow), new object[] { i_Row });
            return;
        }

        this.dataGridView1.Rows.Insert(0, i_Row);
        if (dataGridView1.Rows.Count > Max_Rows_Val)
        {
            dataGridView1.Rows.RemoveAt((dataGridView1.Rows.Count - 1));
        }
        this.dataGridView1.Update();
    }
    bool ExeConnectionDUT()
    {
        AddressIP = textBox1.Text;
        if (AddressIP != "")
        {
            if (HttpReqService.HttpReqTCP_Connet(AddressIP))
            {
                ConnectFlg = true;
                //ExeTimesCnt = 1;                    
            }
            else
            {
                ConnectFlg = false; PrintTitle("DUT Disconnected");
                return false;
            }
        }
        else return false;

        if (ConnectFlg)
        {
            dev = HttpReqService.GetDevice();
            typTxt.Text = devName = dev.ModuleType;
            PrintTitle("DUT [ " + devName + " ] is connecting");
            //
            if (devName == "") return false;
        }

        return true;
    }

    private void button1_Click(object sender, EventArgs e)
    {
        if (!backgroundWorker1.IsBusy)
        {
            if (ExeConnectionDUT())
                this.backgroundWorker1.RunWorkerAsync();
        }
    }

    void PrintTitle(string title)
    {
        DataGridViewRow dgvRow;
        DataGridViewCell dgvCell;
        dgvRow = new DataGridViewRow();
        dgvRow.DefaultCellStyle.Font = new Font(this.Font, FontStyle.Bold);
        dgvCell = new DataGridViewTextBoxCell(); //Column Time
        var dataTimeInfo = DateTime.Now.ToString("yyyy-MM-dd HH:MM:ss");
        dgvCell.Value = dataTimeInfo;
        dgvRow.Cells.Add(dgvCell);
        //
        dgvCell = new DataGridViewTextBoxCell();
        dgvCell.Value = title;
        dgvRow.Cells.Add(dgvCell);
        //
        dgvCell = new DataGridViewTextBoxCell();
        dgvCell.Value = "";
        dgvRow.Cells.Add(dgvCell);
        //
        dgvCell = new DataGridViewTextBoxCell();
        dgvCell.Value = "";
        dgvRow.Cells.Add(dgvCell);

        m_DataGridViewCtrlAddDataRow(dgvRow);
        //Application.DoEvents();
    }
    void PrintStep()
    {
        DataGridViewRow dgvRow;
        DataGridViewCell dgvCell;

        var list = selenium.GetStepResult();
        foreach (var item in list)
        {
            AdvSeleniumAPI.ResultClass _res = (AdvSeleniumAPI.ResultClass)item;
            //
            dgvRow = new DataGridViewRow();
            if (_res.Res == "fail")
            {
                errorCnt++;
                dgvRow.DefaultCellStyle.ForeColor = Color.Red;
            }
            dgvCell = new DataGridViewTextBoxCell(); //Column Time
                                                     //
            if (_res == null) continue;
            //
            var dataTimeInfo = DateTime.Now.ToString("yyyy-MM-dd HH:MM:ss");
            dgvCell.Value = dataTimeInfo;
            dgvRow.Cells.Add(dgvCell);
            //
            dgvCell = new DataGridViewTextBoxCell();
            dgvCell.Value = _res.Decp;
            dgvRow.Cells.Add(dgvCell);
            //
            dgvCell = new DataGridViewTextBoxCell();
            dgvCell.Value = _res.Res;
            dgvRow.Cells.Add(dgvCell);
            //
            dgvCell = new DataGridViewTextBoxCell();
            dgvCell.Value = _res.Tdev;
            dgvRow.Cells.Add(dgvCell);
            //
            dgvCell = new DataGridViewTextBoxCell();
            dgvCell.Value = _res.Err;
            dgvRow.Cells.Add(dgvCell);

            m_DataGridViewCtrlAddDataRow(dgvRow);
        }
        Application.DoEvents();


    }

    //----------------------------------------------------------------------------//
    private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
    {
        //4.在方法中傳遞BackgroundWorker參數
        Running(sender as BackgroundWorker);
    }

    private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
    {
    }

    //----------------------------------------------------------------------------//
    private void Running(BackgroundWorker myWork)//控制訊號源輸出值，並紀錄AI讀值結果
    {
        WorkSteps();
    }

    private void WorkSteps()
    {
        if (!DOModChkBlock())
        {
            PrintTitle("Module not support.");
            return;
        }
        errorCnt = step =0;
        selenium = new AdvSeleniumAPIv2();

        selenium.StartupServer("http://" + textBox1.Text);
        System.Threading.Thread.Sleep(1000);

        selenium.Type("id=ACT0", "root");
        selenium.Type("id=PWD0", "00000000");
        selenium.Click("id=APY0");
        selenium.WaitForPageToLoad("30000");
        PrintStep();
        //
        DOConfgStsBlock();
        selenium.Close();

        
    }
    int do_num = 0;
    bool DOModChkBlock()
    {
        if (dev.ModuleType.ToUpper() == "WISE-4012E" ||
            dev.ModuleType.ToUpper() == "WISE-4012")
        { do_num = 2; return true; }
        else if (dev.ModuleType.ToUpper() == "WISE-4050" 
            || dev.ModuleType.ToUpper() == "WISE-4060"
            || dev.ModuleType.ToUpper() == "WISE-4050/LAN"
            || dev.ModuleType.ToUpper() == "WISE-4060/LAN"
            || dev.ModuleType.ToUpper() == "WISE-4010/LAN")
        { do_num = 4; return true; }
        return false;
    }

    void DOConfgStsBlock()
    {
        PrintTitle("DOConfgSts");
        selenium.Click("id=ioStatus0");

        if (dev.ModuleType.ToUpper() == "WISE-4012")
        {
            selenium.Click("link=DO");
            //selenium.Click("id=doConfigRow0");
            DOExeDOconfig();
            DOExePlsconfig();
            DOExeL2Hdconfig();
            DOExeH2Ldconfig();
            DOExeAlmconfig();
        }
        else if (dev.ModuleType.ToUpper() == "WISE-4012E")
        {
            selenium.Click("link=Relay");
            //selenium.Click("id=doConfigRow0");
            RlyExeDOconfig();
            RlyExePlsconfig();
            RlyExeL2Hdconfig();
            RlyExeH2Ldconfig();
            RlyExeAlmconfig();
        }
        else if (dev.ModuleType.ToUpper() == "WISE-4050"
            || dev.ModuleType.ToUpper() == "WISE-4050/LAN"
            || dev.ModuleType.ToUpper() == "WISE-4010/LAN")
        {
            selenium.Click("link=DO");
            selenium.Click("id=doConfigRow0");
            DOExeDOconfig();
            DOExePlsconfig();
            DOExeL2Hdconfig();
            DOExeH2Ldconfig();
        }
        else if (dev.ModuleType.ToUpper() == "WISE-4060"
            || dev.ModuleType.ToUpper() == "WISE-4060/LAN")
        {
            selenium.Click("link=Relay");
            selenium.Click("id=doConfigRow0");
            RlyExeDOconfig();
            RlyExePlsconfig();
            RlyExeL2Hdconfig();
            RlyExeH2Ldconfig();
        }
        else
            PrintTitle("Module is not support.");
    }

    void EnterDOConfigPage()
    {
        if (dev.ModuleType.ToUpper() == "WISE-4012E" || dev.ModuleType.ToUpper() == "WISE-4012")
            selenium.Click("xpath=(//a[contains(text(),'Configuration')])[3]");
        else
            selenium.Click("xpath=(//a[contains(text(),'Configuration')])[2]");
    }

    void DOExeDOconfig()
    {
        PrintTitle("DOExeDOconfig");
        EnterDOConfigPage();
        for (int i = 0; i < do_num; i++)
        {            
            selenium.Select("css=#doConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-3 > div.input-group > #selCh", "label=" + i.ToString());
            selenium.Type("css=#doConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-6 > div.input-group > #inpTag", "ABCDEFGHIJKLMNOPQRST" + i.ToString());
            selenium.Select("css=#doConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-2 > div.input-group > #selMd", "label=DO");
            selenium.Click("id=inpFSV");
            selenium.Click("css=#doConfigBaseForm > div.panel-footer.clearfix > div.pull-right > #btnSubmit");
            PrintStep();
        }
        //All
        selenium.Select("css=#doConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-3 > div.input-group > #selCh", "label=All");
        selenium.Select("css=#doConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-2 > div.input-group > #selMd", "label=DO");
        selenium.Click("id=inpFSV");
        selenium.Click("css=#doConfigBaseForm > div.panel-footer.clearfix > div.pull-right > #btnSubmit");
        PrintStep();
        // check the view        
        selenium.Click("xpath=(//a[contains(text(),'Status')])[3]");
        for (int i = 0; i < do_num; i++)
            selenium.Click("id=switchDO_" + i.ToString());

    }

    void DOExePlsconfig()
    {
        PrintTitle("DOExePlsconfig");
        EnterDOConfigPage();
        for (int i = 0; i < do_num; i++)
        {            
            selenium.Select("css=#doConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-3 > div.input-group > #selCh", "label=" + i.ToString());
            selenium.Type("css=#doConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-6 > div.input-group > #inpTag", "ABCDEFGHIJKLMNOPQRS" + (i + 10).ToString());
            selenium.Select("css=#doConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-2 > div.input-group > #selMd", "label=Pulse Output");
            selenium.Click("id=inpFSV");
            selenium.Type("id=inpPsLo", "65535");
            selenium.Type("id=inpPsHi", "65535");
            selenium.Click("css=#doConfigBaseForm > div.panel-footer.clearfix > div.pull-right > #btnSubmit");
            PrintStep();
        }
        //All
        selenium.Select("css=#doConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-3 > div.input-group > #selCh", "label=All");
        selenium.Select("css=#doConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-2 > div.input-group > #selMd", "label=Pulse Output");
        selenium.Click("id=inpFSV");
        selenium.Type("id=inpPsLo", "65535");
        selenium.Type("id=inpPsHi", "65535");
        selenium.Click("css=#doConfigBaseForm > div.panel-footer.clearfix > div.pull-right > #btnSubmit");
        PrintStep();
        //check the view
        PrintTitle("Check DO pulse status.");
        selenium.Click("xpath=(//a[contains(text(),'Status')])[3]");
        for (int i = 0; i < do_num; i++)
        {
            selenium.Type("css=#DO_" + i.ToString() + " > #pulseOutConfig > div > div.row > div.input-group > #numbarFixedCount", "4294967295");
            if (dev.ModuleType.ToUpper() == "WISE-4010/LAN")
            {
                selenium.Click("xpath=(//button[@id='psStart'])[" + (i + 1).ToString() + "]");
                selenium.Click("xpath=(//button[@id='psStop'])[" + (i + 1).ToString() + "]");
            }
            else
            {
                selenium.Click("xpath=(//button[@id='psStart'])[" + (i + 5).ToString() + "]");
                selenium.Click("xpath=(//button[@id='psStop'])[" + (i + 5).ToString() + "]");
            }
            PrintStep();
            selenium.Click("name=DoPulseOutModeDO_" + i.ToString());

            if (dev.ModuleType.ToUpper() == "WISE-4010/LAN")
            {
                selenium.Click("xpath=(//button[@id='psStart'])[" + (i + 1).ToString() + "]");
                selenium.Click("xpath=(//button[@id='psStop'])[" + (i + 1).ToString() + "]");
            }
            else
            {
                selenium.Click("xpath=(//button[@id='psStart'])[" + (i + 5).ToString() + "]");
                selenium.Click("xpath=(//button[@id='psStop'])[" + (i + 5).ToString() + "]");
            }
            PrintStep();
        }
    }

    void DOExeL2Hdconfig()
    {
        PrintTitle("DOExeL2Hdconfig");
        EnterDOConfigPage();
        for (int i = 0; i < do_num; i++)
        {
            selenium.Select("css=#doConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-3 > div.input-group > #selCh", "label=" + i.ToString());
            selenium.Type("css=#doConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-6 > div.input-group > #inpTag", "ABCDEFGHIJKLMNOPQRS" + (i + 20).ToString());
            selenium.Select("css=#doConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-2 > div.input-group > #selMd", "label=Low to High Delay");
            selenium.Click("id=inpFSV");
            selenium.Type("id=inpLDT", "65535");
            selenium.Click("css=#doConfigBaseForm > div.panel-footer.clearfix > div.pull-right > #btnSubmit");
            PrintStep();
        }
        selenium.Select("css=#doConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-3 > div.input-group > #selCh", "label=All");
        selenium.Select("css=#doConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-2 > div.input-group > #selMd", "label=Low to High Delay");
        selenium.Click("id=inpFSV");
        selenium.Type("id=inpLDT", "65535");
        selenium.Click("css=#doConfigBaseForm > div.panel-footer.clearfix > div.pull-right > #btnSubmit");
        PrintStep();
        // check the view        
        selenium.Click("xpath=(//a[contains(text(),'Status')])[3]");
        for (int i = 0; i < do_num; i++)
            selenium.Click("id=switchDO_" + i.ToString());
    }

    void DOExeH2Ldconfig()
    {
        PrintTitle("DOExeH2Ldconfig");
        EnterDOConfigPage();
        for (int i = 0; i < do_num; i++)
        {
            selenium.Select("css=#doConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-3 > div.input-group > #selCh", "label=" + i.ToString());
            selenium.Type("css=#doConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-6 > div.input-group > #inpTag", "ABCDEFGHIJKLMNOPQRS" + (i + 30).ToString());
            selenium.Select("css=#doConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-2 > div.input-group > #selMd", "label=High to Low Delay");
            selenium.Click("id=inpFSV");
            selenium.Type("id=inpHDT", "65535");
            selenium.Click("css=#doConfigBaseForm > div.panel-footer.clearfix > div.pull-right > #btnSubmit");
            PrintStep();
        }
        selenium.Select("css=#doConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-3 > div.input-group > #selCh", "label=All");
        selenium.Select("css=#doConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-2 > div.input-group > #selMd", "label=High to Low Delay");
        selenium.Click("id=inpFSV");
        selenium.Type("id=inpHDT", "65535");
        selenium.Click("css=#doConfigBaseForm > div.panel-footer.clearfix > div.pull-right > #btnSubmit");
        PrintStep();
        // check the view        
        selenium.Click("xpath=(//a[contains(text(),'Status')])[3]");
        for (int i = 0; i < do_num; i++)
            selenium.Click("id=switchDO_" + i.ToString());
        //
        step = 99;
        PrintTitle("Step is end.");
    }

    void DOExeAlmconfig()
    {
        PrintTitle("DOExeAlmconfigFor4012");
        EnterDOConfigPage();
        for (int i = 0; i < do_num; i++)
        {
            selenium.Select("css=#doConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-3 > div.input-group > #selCh", "label=" + i.ToString());
            selenium.Select("css=#doConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-2 > div.input-group > #selMd", "label=AI Alarm Driven");
            selenium.Select("id=selAMd", "label=High Alarm");
            selenium.Select("id=selACh", "label=" + (2 - i).ToString());
            selenium.Click("css=#doConfigBaseForm > div.panel-footer.clearfix > div.pull-right > #btnSubmit");
            PrintStep();
        }
        selenium.Select("css=#doConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-3 > div.input-group > #selCh", "label=All");
        selenium.Select("css=#doConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-2 > div.input-group > #selMd", "label=AI Alarm Driven");
        selenium.Select("id=selAMd", "label=Low Alarm");
        selenium.Select("id=selACh", "label=2");
        selenium.Click("css=#doConfigBaseForm > div.panel-footer.clearfix > div.pull-right > #btnSubmit");
        PrintStep();
        //
        step = 99;
        PrintTitle("Step is end.");
    }

    //=======================================================//
    void RlyExeDOconfig()
    {
        PrintTitle("ExeRelayConfig");
        EnterDOConfigPage();
        for (int i = 0; i < do_num; i++)
        {
            selenium.Select("css=#relayConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-3 > div.input-group > #selCh", "label=" + i.ToString());
            selenium.Type("css=#relayConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-6 > div.input-group > #inpTag", "ABCDEFGHIJKLMNOPQRST" + i.ToString());
            selenium.Select("css=#relayConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-2 > div.input-group > #selMd", "label=DO");
            selenium.Click("id=inpFSV");
            selenium.Click("css=#relayConfigBaseForm > div.panel-footer.clearfix > div.pull-right > #btnSubmit");
            PrintStep();
        }
        //All
        selenium.Select("css=#relayConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-3 > div.input-group > #selCh", "label=All");
        selenium.Select("css=#relayConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-2 > div.input-group > #selMd", "label=DO");
        selenium.Click("id=inpFSV");
        selenium.Click("css=#relayConfigBaseForm > div.panel-footer.clearfix > div.pull-right > #btnSubmit");
        PrintStep();
        // check the view        
        selenium.Click("xpath=(//a[contains(text(),'Status')])[3]");
        for (int i = 0; i < do_num; i++)
            selenium.Click("id=switchDO_" + i.ToString());
    }

    void RlyExePlsconfig()
    {
        PrintTitle("ExeRelayPlsConfig");
        EnterDOConfigPage();
        for (int i = 0; i < do_num; i++)
        {
            selenium.Select("css=#relayConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-3 > div.input-group > #selCh", "label=" + i.ToString());
            selenium.Type("css=#relayConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-6 > div.input-group > #inpTag", "ABCDEFGHIJKLMNOPQRS" + (i + 10).ToString());
            selenium.Select("css=#relayConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-2 > div.input-group > #selMd", "label=Pulse Output");
            selenium.Click("id=inpFSV");
            selenium.Type("id=inpPsLo", "65535");
            selenium.Type("id=inpPsHi", "65535");
            selenium.Click("css=#relayConfigBaseForm > div.panel-footer.clearfix > div.pull-right > #btnSubmit");
            PrintStep();
        }
        //All
        selenium.Select("css=#relayConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-3 > div.input-group > #selCh", "label=All");
        selenium.Select("css=#relayConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-2 > div.input-group > #selMd", "label=Pulse Output");
        selenium.Click("id=inpFSV");
        selenium.Type("id=inpPsLo", "65535");
        selenium.Type("id=inpPsHi", "65535");
        selenium.Click("css=#relayConfigBaseForm > div.panel-footer.clearfix > div.pull-right > #btnSubmit");
        PrintStep();
        //check the view
        if (dev.ModuleType.ToUpper() == "WISE-4012E")
            selenium.Click("xpath=(//a[contains(text(),'Status')])[4]");
        else
            selenium.Click("xpath=(//a[contains(text(),'Status')])[3]");
        for (int i = 0; i < do_num; i++)
        {
            if (dev.ModuleType.ToUpper() == "WISE-4012E")
            {
                selenium.Type("css=#DO_" + i.ToString() + "> #pulseOutConfig > div > div.row > div.input-group > #numbarFixedCount", "4294967295");
                selenium.Click("xpath=(//button[@id='psStart'])[" + (i + 3).ToString() + "]");
                selenium.Click("xpath=(//button[@id='psStop'])[" + (i + 3).ToString() + "]");
                selenium.Click("name=DoPulseOutModeDO_" + i.ToString());
                selenium.Click("xpath=(//button[@id='psStart'])[" + (i + 3).ToString() + "]");
                selenium.Click("xpath=(//button[@id='psStop'])[" + (i + 3).ToString() + "]");
                //selenium.Click("xpath=(//button[@id='psStart'])[4]");
                //selenium.Click("xpath=(//button[@id='psStop'])[4]");
                //selenium.Click("name=DoPulseOutModeDO_1");
                //selenium.Click("xpath=(//button[@id='psStart'])[4]");
                //selenium.Click("xpath=(//button[@id='psStop'])[4]");
            }
            else
            {
                selenium.Type("css=#DO_" + i.ToString() + " > #pulseOutConfig > div > div.row > div.input-group > #numbarFixedCount", "4294967295");
                selenium.Click("xpath=(//button[@id='psStart'])[" + (i + 5).ToString() + "]");
                selenium.Click("xpath=(//button[@id='psStop'])[" + (i + 5).ToString() + "]");
                selenium.Click("name=DoPulseOutModeDO_" + i.ToString());
                selenium.Click("xpath=(//button[@id='psStart'])[" + (i + 5).ToString() + "]");
                selenium.Click("xpath=(//button[@id='psStop'])[" + (i + 5).ToString() + "]");
            }
                
            PrintStep();
        }
    }

    void RlyExeL2Hdconfig()
    {
        PrintTitle("ExeRelayeL2HConfig");
        EnterDOConfigPage();
        for (int i = 0; i < do_num; i++)
        {
            selenium.Select("css=#relayConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-3 > div.input-group > #selCh", "label=" + i.ToString());
            selenium.Type("css=#relayConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-6 > div.input-group > #inpTag", "ABCDEFGHIJKLMNOPQRS" + (i + 20).ToString());
            selenium.Select("css=#relayConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-2 > div.input-group > #selMd", "label=Low to High Delay");
            selenium.Click("id=inpFSV");
            selenium.Type("id=inpLDT", "65535");
            selenium.Click("css=#relayConfigBaseForm > div.panel-footer.clearfix > div.pull-right > #btnSubmit");
            PrintStep();
        }
        selenium.Select("css=#relayConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-3 > div.input-group > #selCh", "label=All");
        selenium.Select("css=#relayConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-2 > div.input-group > #selMd", "label=Low to High Delay");
        selenium.Click("id=inpFSV");
        selenium.Type("id=inpLDT", "65535");
        selenium.Click("css=#relayConfigBaseForm > div.panel-footer.clearfix > div.pull-right > #btnSubmit");
        PrintStep();
        // check the view        
        selenium.Click("xpath=(//a[contains(text(),'Status')])[3]");
        for (int i = 0; i < do_num; i++)
            selenium.Click("id=switchDO_" + i.ToString());
    }

    void RlyExeH2Ldconfig()
    {
        PrintTitle("ExeRelayeH2LConfig");
        EnterDOConfigPage();
        for (int i = 0; i < do_num; i++)
        {
            selenium.Select("css=#relayConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-3 > div.input-group > #selCh", "label=" + i.ToString());
            selenium.Type("css=#relayConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-6 > div.input-group > #inpTag", "ABCDEFGHIJKLMNOPQRS" + (i + 30).ToString());
            selenium.Select("css=#relayConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-2 > div.input-group > #selMd", "label=High to Low Delay");
            selenium.Click("id=inpFSV");
            selenium.Type("id=inpHDT", "65535");
            selenium.Click("css=#relayConfigBaseForm > div.panel-footer.clearfix > div.pull-right > #btnSubmit");
            PrintStep();
        }
        selenium.Select("css=#relayConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-3 > div.input-group > #selCh", "label=All");
        selenium.Select("css=#relayConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-2 > div.input-group > #selMd", "label=High to Low Delay");
        selenium.Click("id=inpFSV");
        selenium.Type("id=inpHDT", "65535");
        selenium.Click("css=#relayConfigBaseForm > div.panel-footer.clearfix > div.pull-right > #btnSubmit");
        PrintStep();
        // check the view        
        selenium.Click("xpath=(//a[contains(text(),'Status')])[3]");
        for (int i = 0; i < do_num; i++)
            selenium.Click("id=switchDO_" + i.ToString());
        //
        step = 99;
        PrintTitle("Step is end.");
    }
    void RlyExeAlmconfig()
    {
        PrintTitle("RelayExeAlmconfigFor4012E");
        EnterDOConfigPage();
        for (int i = 0; i < do_num; i++)
        {
            selenium.Select("css=#relayConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-3 > div.input-group > #selCh", "label=" + i.ToString());
            selenium.Select("css=#relayConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-2 > div.input-group > #selMd", "label=AI Alarm Driven");
            selenium.Select("id=selAMd", "label=High Alarm");
            selenium.Select("id=selACh", "label=" + (1 - i).ToString());
            selenium.Click("css=#relayConfigBaseForm > div.panel-footer.clearfix > div.pull-right > #btnSubmit");
            PrintStep();
        }
        selenium.Select("css=#relayConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-3 > div.input-group > #selCh", "label=All");
        selenium.Select("css=#relayConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-2 > div.input-group > #selMd", "label=AI Alarm Driven");
        selenium.Select("id=selAMd", "label=Low Alarm");
        selenium.Select("id=selACh", "label=1");
        selenium.Click("css=#relayConfigBaseForm > div.panel-footer.clearfix > div.pull-right > #btnSubmit");
        PrintStep();
        //
        step = 99;
        PrintTitle("Step is end.");
    }
}//class
