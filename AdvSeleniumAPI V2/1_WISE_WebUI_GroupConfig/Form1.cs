using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using AdvSeleniumAPI;
using iATester;
using Model;
using Service;
using ThirdPartyToolControl;

public partial class Form1 : Form, iATester.iCom
{
    private delegate void DataGridViewCtrlAddDataRow(DataGridViewRow i_Row);
    private DataGridViewCtrlAddDataRow m_DataGridViewCtrlAddDataRow;
    internal const int Max_Rows_Val = 65535;

    IHttpReqService HttpReqService;
    DataHandleService dataHld;
    DeviceModel dev; IniDataFmt LoadPara;
    string AddressIP = "", devName = "", path = "", browser = ""; bool ConnectFlg = false;
    int errorCnt = 0;
    AdvSeleniumAPIv2 selenium;
    cThirdPartyToolControl tpc = new cThirdPartyToolControl();
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
        LoadPara = dataHld.GetPara(Application.StartupPath);
        textBox1.Text = AddressIP = LoadPara.IP;
        GCtxtBox.Text = path = LoadPara.Path;
    }
    public void StartTest()//iATester
    {
        if (ExeConnectionDUT())
        {
            eStatus(this, new StatusEventArgs(iStatus.Running));
            WorkSteps();
            if(errorCnt > 0)
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
            //save para.
            LoadPara.IP = AddressIP;
            LoadPara.Path = GCtxtBox.Text;
            dataHld.SavePara("", LoadPara);
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
    private void GCBtn_Click(object sender, EventArgs e)
    {
        DialogResult result = folderBrowserDialog1.ShowDialog();
        if (result == DialogResult.OK)
        {
            GCtxtBox.Text = path = folderBrowserDialog1.SelectedPath;
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
        PrintTitle("Check folder path.");
        errorCnt = 0;
        string filepath = GCtxtBox.Text;
        if (filepath == "" || filepath == null)
        {
            PrintTitle("Path failed!");
            return;
        }

        selenium = new AdvSeleniumAPIv2();

        selenium.StartupServer("http://" + textBox1.Text);
        System.Threading.Thread.Sleep(1000);

        selenium.Type("id=ACT0", "root");
        selenium.Type("id=PWD0", "00000000");
        selenium.Click("id=APY0");
        selenium.WaitForPageToLoad("30000");
        PrintStep();
        //
        selenium.Click("id=configuration");
        selenium.Click("link=Firmware");
        selenium.Click("id=groupConfigIpSettingBtn");
        selenium.Click("link=With IP Settings");
        selenium.Click("id=inpGroupConfig");
        PrintStep();
        //
        String pathToFile = "";
        if (dev.ModuleType.ToUpper() == "WISE-4012E")
            pathToFile = @filepath + "\\config_file_wise4012e.cfg";
        else if (dev.ModuleType.ToUpper() == "WISE-4012")
            pathToFile = @filepath + "\\config_file_wise4012.cfg";
        else if (dev.ModuleType.ToUpper() == "WISE-4051")
            pathToFile = @filepath + "\\config_file_wise4051.cfg";
        else if (dev.ModuleType.ToUpper() == "WISE-4050")
            pathToFile = @filepath + "\\config_file_wise4050.cfg";
        else if (dev.ModuleType.ToUpper() == "WISE-4060")
            pathToFile = @filepath + "\\config_file_wise4060.cfg";
        else if (dev.ModuleType.ToUpper() == "WISE-4050/LAN")
            pathToFile = @filepath + "\\config_file_wise4050lan.cfg";
        else if (dev.ModuleType.ToUpper() == "WISE-4060/LAN")
            pathToFile = @filepath + "\\config_file_wise4060lan.cfg";
        else if (dev.ModuleType.ToUpper() == "WISE-4010/LAN")
            pathToFile = @filepath + "\\config_file_wise4010lan.cfg";

        //
        System.Threading.Thread.Sleep(3000); int Main_Handl = 0;
        int iLoginKeyboard_Handle = tpc.F_FindWindow("#32770", "上傳檔案");
        int iIE_Handl_1 = tpc.F_FindWindowEx(iLoginKeyboard_Handle, 0, "ComboBoxEx32", "");
        int iIE_Handl_2 = tpc.F_FindWindowEx(iIE_Handl_1, 0, "ComboBox", "");
        Main_Handl = tpc.F_FindWindowEx(iIE_Handl_2, 0, "Edit", "");

        if (Main_Handl > 0)
        {
            System.Threading.Thread.Sleep(1000);
            SendCharToHandle(Main_Handl, 1, pathToFile);
            System.Threading.Thread.Sleep(1000);
            tpc.F_PostMessage(Main_Handl, tpc.V_WM_KEYDOWN, tpc.V_VK_RETURN, 0);
            System.Threading.Thread.Sleep(1000);
            //
            selenium.Click("//a[@id='btnGroupConfig']/i");
            PrintStep();
            System.Threading.Thread.Sleep(20000);//delay 20sec
        }
        else PrintTitle("Get Handle Fail."); 
        //MessageBox.Show("Get Handle Fail.");
        //
        //System.Threading.Thread.Sleep(3000);
        //SendKeys.SendWait(pathToFile);
        //System.Threading.Thread.Sleep(3000);
        //SendKeys.SendWait("{ENTER}");
        
        PrintTitle("Browser close.");
        selenium.Close();
    }
    private void SendCharToHandle(int iHandle, int iDelay, string sText)
    {
        var chars = sText.ToCharArray();
        for (int ctr = 0; ctr < chars.Length; ctr++)
        {
            tpc.F_PostMessage(iHandle, tpc.V_WM_CHAR, chars[ctr], 0);
            System.Threading.Thread.Sleep(iDelay);
        }
    }
}//class
