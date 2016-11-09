using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Net;
using System.Text;
using System.Windows.Forms;
using System.IO;

using AdvWebUIAPI;
using AutoWebUI_ClassLibrary;
using Model;
using Service;
using iATester;
using ThirdPartyToolControl;

/// <summary>
/// Disable WISE IO Access Tag and check cloud would get the tag was deleted.
/// </summary>
public partial class Form1 : Form, iATester.iCom
{
    IAdvSeleniumAPI api;
    private delegate void DataGridViewCtrlAddDataRow(DataGridViewRow i_Row);
    private DataGridViewCtrlAddDataRow m_DataGridViewCtrlAddDataRow;
    internal const int Max_Rows_Val = 65535;

    IHttpReqService HttpReqService;
    DeviceModel dev;
    int errorCnt = 0;
    string AddressIP = "", devName = "", path = "", browser = ""; bool ConnectFlg = false;
    string filename = "WISE_MQTT_CONFIG.ini";
    string folderPath = "";
    int ai_point = 0, di_point = 0, do_point = 0;
    cThirdPartyToolControl tpc = new cThirdPartyToolControl();

    //iATester
    //Send Log data to iAtester
    public event EventHandler<LogEventArgs> eLog = delegate { };
    //Send test result to iAtester
    public event EventHandler<ResultEventArgs> eResult = delegate { };
    //Send execution status to iAtester
    public event EventHandler<StatusEventArgs> eStatus = delegate { };

    public Form1()
    {
        InitializeComponent();
    }

    private void Form1_Load(object sender, EventArgs e)
    {
        HttpReqService = new HttpReqService();
        //
        //dataGridView1.ColumnCount = 1;
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
        backgroundWorker1.WorkerSupportsCancellation = true;
        GetParaFromFile();
    }

    private void Form1_FormClosed(object sender, FormClosedEventArgs e)
    {
        if (backgroundWorker1.IsBusy) backgroundWorker1.CancelAsync();
    }
    public void StartTest()//iATester
    {
        GetParaFromFile();
        if (ExeConnectionDUT())
        {
            eStatus(this, new StatusEventArgs(iStatus.Running));
            WorkSteps();
            if (errorCnt > 0)
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

        if (ConnectFlg)
        {
            dev = HttpReqService.GetDevice();
            devName = dev.ModuleType;
            PrintTitle("DUT [ " + devName + " ] is connecting");
            SetParaToFile();
            //
            if (devName == "") return false;
            CheckModPoint();
        }

        return true;
    }
    void GetParaFromFile()
    {
        string sPath = System.Reflection.Assembly.GetAssembly(this.GetType()).Location;
        char delimiterChars = '\\';
        string[] words = sPath.Split(delimiterChars);
        folderPath = "";
        for (int i = 0; i < words.Length - 1; i++)
        {
            folderPath = folderPath + words[i] + "\\";
        }

        if (File.Exists(folderPath + "\\" + filename))
        {
            using (ExecuteIniClass IniFile = new ExecuteIniClass(Path.Combine(folderPath, filename)))
            {
                textBox1.Text = IniFile.getKeyValue("Dev", "IP");
                txtCloudIp.Text = IniFile.getKeyValue("Dev", "Cloud");
            }
        }
    }
    void SetParaToFile()
    {
        if (!File.Exists(folderPath + "\\" + filename))
            File.Create(folderPath + "\\" + filename);

        //save para.
        using (ExecuteIniClass IniFile = new ExecuteIniClass(Path.Combine(folderPath, filename)))
        {
            IniFile.setKeyValue("Dev", "IP", textBox1.Text);
            IniFile.setKeyValue("Dev", "Cloud", txtCloudIp.Text);
        }
    }
    private void button1_Click(object sender, EventArgs e)
    {
        if (!backgroundWorker1.IsBusy)
        {
            if (ExeConnectionDUT())
                this.backgroundWorker1.RunWorkerAsync();
        }
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
        errorCnt = 0;
        api = new AdvSeleniumAPI("FireFox", Application.StartupPath);
        System.Threading.Thread.Sleep(1000);
        LinkWebBlock();
        LogInBlock();
        CheckWAConfig();
        CloseBrowserBlock();
        System.Threading.Thread.Sleep(1000);
        //<------------------ 確認雲端上刪除Tag的部份
        api = new AdvSeleniumAPI("IE", Application.StartupPath);
        System.Threading.Thread.Sleep(1000);
        ViewandCheckCloudTagInfo();

        //<------------------
        CloseBrowserBlock();
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
    }

    void PrintStep()
    {
        DataGridViewRow dgvRow;
        DataGridViewCell dgvCell;

        var list = api.GetStepResult();
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
    }
    //
    void CloseBrowserBlock()
    {
        PrintTitle("CloseBrowser");
        if (api != null) api.Quit();
    }

    void LinkWebBlock()
    {
        api.LinkWebUI("http://" + AddressIP + "/config");
        if (browser == "FireFox")
        { System.Threading.Thread.Sleep(1000); api.ZoomWebUI(); }
        System.Threading.Thread.Sleep(1000);
    }

    void LogInBlock()
    {
        PrintTitle("LogIn");
        api.Enter("root").ById("ACT0").Exe();
        api.Enter("00000000").ById("PWD0").Exe();
        api.ById("APY0").ClickAndWait(1000);
        PrintStep();
    }

    void CheckWAConfig()
    {
        PrintTitle("Check Enable Status");
        api.ById("advancedFunction").ClickAndWait(1000);
        api.ById("dataLog").ClickAndWait(1000);
        api.ByTxt("WebAccess I/O Configuration").ClickAndWait(1000);
        //
        PrintTitle("Disable IO Tag");
        if (di_point > 0)
        {
            if (api.ByXpath("(//input[@type='checkbox'])[4]").GetAttr("checked") == "true")
            {
                PrintTitle("Disable DI-0 checkbox");
                api.ByXpath("(//input[@type='checkbox'])[4]").Click();
                api.ByXpath("//div[@id='tabWebAccessLog']/div/div/div/div[2]/div/button").Click();
            }
        }

        if (do_point > 0)
        {
            api.ByTxt("DO/Relay").ClickAndWait(1000);
            if (api.ByXpath("(//input[@type='checkbox'])[14]").GetAttr("checked") == "true")
            {
                PrintTitle("Disable DO-0 checkbox");
                api.ByXpath("(//input[@type='checkbox'])[14]").Click();
                api.ByXpath("//div[@id='tabWebAccessLog']/div/div/div/div[2]/div/button").Click();
            }
        }

        if (ai_point > 0)
        {
            api.ByTxt("AI").ClickAndWait(1000);
            if (api.ByXpath("(//input[@type='checkbox'])[20]").GetAttr("checked") == "true")
            {
                PrintTitle("Disable AI-0 checkbox");
                api.ByXpath("(//input[@type='checkbox'])[20]").Click();
                api.ByXpath("//div[@id='tabWebAccessLog']/div/div/div/div[2]/div/button").Click();
            }
        }
        PrintStep();
    }

    void CheckModPoint()
    {
        if (dev.ModuleType.ToUpper() == "WISE-4012E")
        {
            ai_point = 2; di_point = 2; do_point = 2;
        }
        else if (dev.ModuleType.ToUpper() == "WISE-4012"
            || dev.ModuleType.ToUpper() == "WISE-4010/LAN")
        {
            ai_point = 4; di_point = 0; do_point = 2;
        }
        else if (dev.ModuleType.ToUpper() == "WISE-4051")
        {
            ai_point = 0; di_point = 8; do_point = 0;
        }
        else
        {
            ai_point = 0; di_point = 4; do_point = 4;
        }
    }

    string RenameModname()
    {
        string _name = dev.ModuleType;
        if (dev.ModuleType.ToUpper() == "WISE-4050/LAN"
            || dev.ModuleType.ToUpper() == "WISE-4060/LAN"
            || dev.ModuleType.ToUpper() == "WISE-4010/LAN")
        {
            char delimiterChars = '/';
            string[] words = _name.Split(delimiterChars);
            _name = "";
            foreach (string _w in words)
                _name = _name + _w;
        }
        return _name;
    }

    private void ViewandCheckCloudTagInfo()
    {
        PrintTitle("ViewandCheckCloudTagInfo");
        api.LinkWebUI(txtCloudIp.Text + "/broadWeb/bwconfig.asp?username=admin");
        api.ById("userField").Enter("").Submit().Exe();
        PrintStep();

        // Configure project by project name
        string sProjectName = "WISE%2DDQA"; // WISE-DQA
        api.ByXpath("//a[contains(@href, '/broadWeb/bwMain.asp?pos=project') and contains(@href, 'ProjName=" 
                        + sProjectName + "')]").Click();
        PrintStep();

        api.SwitchToCurWindow(0);
        api.SwitchToFrame("leftFrame", 0);

        api.ByXpath("//td[2]/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td/a/font").Click(); 
        // 檢查AUTOTEST-DO0是否可以正確被選取到 若刪除AUTOTEST-AI0 tag失敗，那會取到AUTOTEST-AI0的值
        System.Threading.Thread.Sleep(2000);

        api.SwitchToCurWindow(0);
        api.SwitchToFrame("rightFrame", 0);
        if (ai_point > 0)
        {
            PrintTitle("Check the next tag name AI-1");
            string sTagChangedName = api.ByXpath("//tr[2]/td[2]").GetText().Trim();
            if (sTagChangedName != "AUTOTEST-AI_1")
                PrintTitle("Fail for result [" + sTagChangedName + "].");
            else
                PrintTitle("Success");
        }
        else if (di_point > 0)
        {
            PrintTitle("Check the next tag name DI-1");
            string sTagChangedName = api.ByXpath("//tr[2]/td[2]").GetText().Trim();
            if (sTagChangedName != "AUTOTEST-DI_1")
                PrintTitle("Fail for result [" + sTagChangedName + "].");
            else
                PrintTitle("Success");
        }
        PrintStep();

    }   // G2C




}
