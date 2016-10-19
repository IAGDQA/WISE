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

public partial class Form1 : Form, iATester.iCom
{
    IAdvSeleniumAPI api;
    private delegate void DataGridViewCtrlAddDataRow(DataGridViewRow i_Row);
    private DataGridViewCtrlAddDataRow m_DataGridViewCtrlAddDataRow;
    internal const int Max_Rows_Val = 65535;

    IHttpReqService HttpReqService;
    DeviceModel dev;
    string AddressIP = "", devName = "", path = "", browser = ""; bool ConnectFlg = false;
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
    }

    private void Form1_FormClosed(object sender, FormClosedEventArgs e)
    {
        if (backgroundWorker1.IsBusy) backgroundWorker1.CancelAsync();
    }
    public void StartTest()//iATester
    {
        //if (ExeConnectionDUT())
        //{
        //    eStatus(this, new StatusEventArgs(iStatus.Running));
        //    WorkSteps();
        //    eResult(this, new ResultEventArgs(iResult.Pass));
        //}
        //else
        //    eResult(this, new ResultEventArgs(iResult.Fail));
        ////
        //eStatus(this, new StatusEventArgs(iStatus.Completion));
        //Application.DoEvents();
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
        GetPara();
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
            //
            if (devName == "") return false;
            CheckModPoint();
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
        api = new AdvSeleniumAPI("IE", Application.StartupPath);
        System.Threading.Thread.Sleep(1000);
        LinkWebBlock();
        LogInBlock();
        ChangeIO_TagName();
        CloudConfig();
        System.Threading.Thread.Sleep(5000);
        //<------------------ 確認雲端上的部份
        ViewandSaveCloudTagInfo();


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
                dgvRow.DefaultCellStyle.ForeColor = Color.Red;
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
    string filename = "WISE_WEB_CONFIG.ini";
    void GetPara()
    {
        //create new
        if (File.Exists(Application.StartupPath + "\\" + filename))
        {
            using (ExecuteIniClass IniFile = new ExecuteIniClass(Path.Combine(Application.StartupPath, filename)))
            {
                AddressIP = IniFile.getKeyValue("Dev", "IP");
                path = IniFile.getKeyValue("Dev", "Path");
                browser = IniFile.getKeyValue("Dev", "Browser");
                PrintTitle("Get Dev: [IP] " + AddressIP
                           + " ; [path] " + path);
            }
        }
        else
            PrintTitle("Get file fail...");

    }

    //
    void CloseBrowserBlock()
    {
        PrintTitle("CloseBrowser");
        if (api != null) api.Quit();
    }

    void LinkWebBlock()
    {
        api.LinkWebUI("http://" + AddressIP);
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

    void ChangeIO_TagName()
    {
        if (di_point > 0)
            ChangeDIConfig();

        if (do_point > 0)
            ChangeDOConfig();

        if (ai_point > 0)
            ChangeAIConfig();
    }

    void CloudConfig()
    {
        api.ById("configuration").ClickAndWait(1000);
        api.ByTxt("Cloud").ClickAndWait(3000);
        api.Select(4).ByXpath("//select[@id='selCloud']").Exe();
        api.Wait(1000);
        api.Enter(txtCloudIp.Text).Clear().ById("Nm").Exe();
        api.Enter("WISE-DQA").Clear().ById("PNm").Exe();
        api.Enter(dev.ModuleType + "-AUTOTEST").Clear().ById("NNm").Exe();
        api.Enter("60").Clear().ById("HbF").Exe();
        api.Enter("80").Clear().ById("PWeb").Exe();
        api.Enter("admin").Clear().ById("Pu").Exe();
        api.Enter("12345").Clear().ById("Pw").Exe();
        api.ById("btnWebAccessSubmit").Click();
        PrintStep();
    }

    void ChangeDIConfig()
    {
        PrintTitle("DIConfgSts");
        api.ById("ioStatus0").ClickAndWait(1000);
        api.ByTxt("DI").ClickAndWait(1000);
        EnterDIConfigPage();
        //
        for (int i = 0; i < di_point; i++)
        {
            api.SelectTxt(i.ToString()).ByCss("#diConfigBaseForm > div.panel-heading "
                + "> div.form-group > div.col-lg-3 > div.input-group > #selCh").Exe();
            api.Enter("AUTOTEST-" + "DI_" + i.ToString()).Clear().ByCss("#diConfigBaseForm > div.panel-heading "
                + "> div.form-group > div.col-lg-6 > div.input-group > #inpTag").Exe();

            api.SelectTxt("DI").ByCss("#selMd").Exe();
            api.ByCss("#diConfigBaseForm > div.panel-footer.clearfix > div.pull-right > #btnSubmit")
                    .ClickAndWait(1000);
        }
    }

    void ChangeDOConfig()
    {
        PrintTitle("DOConfgSts");
        api.ById("ioStatus0").ClickAndWait(1000);
        if (dev.ModuleType.ToUpper() == "WISE-4012"
            || dev.ModuleType.ToUpper() == "WISE-4050")
        {
            api.ByTxt("DO").ClickAndWait(1000);
            DOExeDOconfig();
        }
        else if (dev.ModuleType.ToUpper() == "WISE-4012E"
            || dev.ModuleType.ToUpper() == "WISE-4060")
        {
            api.ByTxt("Relay").ClickAndWait(1000);
            RlyExeDOconfig();
        }
    }

    void ChangeAIConfig()
    {
        UIConfgBlock();
        AIConfgChnBlock();
    }

    void DOExeDOconfig()
    {
        PrintTitle("DOExeDOconfig4012");
        EnterDOConfigPage();
        for (int i = 0; i < do_point; i++)
        {
            api.SelectTxt(i.ToString()).ByCss("#doConfigBaseForm > div.panel-heading > div.form-group " +
                                    "> div.col-lg-3 > div.input-group > #selCh").Exe();
            api.Wait(1000);
            api.Enter("AUTOTEST-" + "DO_" + i.ToString()).Clear().ByCss("#doConfigBaseForm > div.panel-heading "
                + "> div.form-group > div.col-lg-6 > div.input-group > #inpTag").Exe();

            api.SelectTxt("DO").ByCss("#doConfigBaseForm > div.panel-heading > div.form-group " +
                                    "> div.col-lg-2 > div.input-group > #selMd").Exe();

            api.ByCss("#doConfigBaseForm > div.panel-footer.clearfix > div.pull-right > #btnSubmit")
                        .ClickAndWait(1000);
        }
        PrintStep();
    }

    void RlyExeDOconfig()
    {
        PrintTitle("DOExeDOconfig4012e");
        EnterDOConfigPage();
        for (int i = 0; i < do_point; i++)
        {            
            api.SelectTxt(i.ToString()).ByCss("#relayConfigBaseForm > div.panel-heading > div.form-group " +
                                    "> div.col-lg-3 > div.input-group > #selCh").Exe();
            api.Wait(1000);
            api.Enter("AUTOTEST-" + "DO_" + i.ToString()).Clear().ByCss("#relayConfigBaseForm > div.panel-heading "
                + "> div.form-group > div.col-lg-6 > div.input-group > #inpTag").Exe();

            api.SelectTxt("DO").ByCss("#relayConfigBaseForm > div.panel-heading > div.form-group " +
                                    "> div.col-lg-2 > div.input-group > #selMd").Exe();

            api.ByCss("#relayConfigBaseForm > div.panel-footer.clearfix > div.pull-right > #btnSubmit")
                        .ClickAndWait(1000);
        }
        PrintStep();
    }

    void UIConfgBlock()
    {
        if (dev.ModuleType.ToUpper() == "WISE-4012")
        {
            PrintTitle("UIConfg");
            api.ById("ioStatus0").ClickAndWait(1000);
            //
            for (int i = 0; i < 4; i++)
            {
                if (api.ById("inpEn_" + i.ToString()).GetAttr("checked") == "false")
                    api.ById("inpEn_" + i.ToString()).Click();
                api.SelectTxt("AI").ByCss("#selMd_" + i.ToString()).Exe();
            }
            api.ById("btnUIConfig").ClickAndWait(1000);
            PrintStep();
        }
    }

    void AIConfgChnBlock()
    {
        PrintTitle("AIConfgChn");
        api.ById("ioStatus0").ClickAndWait(1000);
        api.ByTxt("AI").ClickAndWait(1000);
        api.ByXpath("//a[contains(text(),'Configuration')]").ClickAndWait(1000);
        //
        if (dev.ModuleType.ToUpper() == "WISE-4012")
        {
            Exe4012_AISetting();
        }
        else if (dev.ModuleType.ToUpper() == "WISE-4012E")
        {
            Exe4012E_AISetting();
        }

        PrintStep();
    }

    void Exe4012_AISetting()
    {
        PrintTitle("Exe4012_AISetting");
        for (int i = 0; i < ai_point; i++)
        {
            api.SelectTxt(i.ToString()).ByCss("div.input-group > #selCh").Exe();
            api.Enter("AUTOTEST-" + "AI_" + i.ToString()).Clear().ById("inpTag").Exe();
            api.ByCss("#aiConfigBaseForm > div.panel-footer.clearfix > div.pull-right > #btnSubmit")
                .ClickAndWait(1000);
            PrintStep();
        }
    }

    void Exe4012E_AISetting()
    {
        PrintTitle("Exe4012E_AISetting");
        for (int i = 0; i < ai_point; i++)
        {
            api.SelectTxt(i.ToString()).ByCss("div.input-group > #selCh").Exe();
            api.Enter("AUTOTEST-" + "AI_" + i.ToString()).Clear().ById("inpTag").Exe();
            api.ByCss("#aiConfigBaseForm > div.panel-footer.clearfix > div.pull-right > #btnSubmit")
                .ClickAndWait(1000);
            PrintStep();
        }
    }

    void CheckModPoint()
    {
        if (dev.ModuleType.ToUpper() == "WISE-4012E")
        {
            ai_point = 2; di_point = 2; do_point = 2;
        }
        else if (dev.ModuleType.ToUpper() == "WISE-4012")
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

    void EnterDIConfigPage()
    {
        if (dev.ModuleType.ToUpper() == "WISE-4012E" || dev.ModuleType.ToUpper() == "WISE-4012")
            api.ByXpath("(//a[contains(text(),'Configuration')])[2]").ClickAndWait(1000);
        else
            api.ByXpath("//a[contains(text(),'Configuration')]").ClickAndWait(1000);
    }

    void EnterDOConfigPage()
    {
        if (dev.ModuleType.ToUpper() == "WISE-4012E" || dev.ModuleType.ToUpper() == "WISE-4012")
            api.ByXpath("(//a[contains(text(),'Configuration')])[3]").ClickAndWait(1000);
        else
            api.ByXpath("(//a[contains(text(),'Configuration')])[2]").ClickAndWait(1000);
    }

    private void ViewandSaveCloudTagInfo()
    {
        PrintTitle("Start WA View");
        api.LinkWebUI(txtCloudIp.Text + "/broadWeb/bwconfig.asp?username=admin");
        api.ById("userField").Enter("").Submit().Exe();
        PrintStep();
        // Configure project by project name

        string sProjectName = "WISE%2DDQA"; // WISE-DQA
        api.ByXpath("//a[contains(@href, '/broadWeb/bwMain.asp?pos=project') and contains(@href, 'ProjName=" + sProjectName + "')]").Click();
        PrintStep();

        // Start kernel
        PrintTitle("Start kernel");
        StartNode();

        System.Threading.Thread.Sleep(5000);

        // Start view
        PrintTitle("Start view");
        api.SwitchToCurWindow(0);
        api.SwitchToFrame("rightFrame", 0);
        api.ByXpath("//tr[2]/td/a/font").Click();
        PrintStep();

        System.Threading.Thread.Sleep(1000);
        // Control browser
        //string sModuleType = "WISE-4012";
        PrintTitle("Control browser");
        int iIE_Handl = tpc.F_FindWindow("IEFrame", "Node : " + dev.ModuleType + "-AUTOTEST - main:untitled");   // 注意是CTestSCADA而不是TestSCADA!! Jammy這邊要改!!
        int iIE_Handl_2 = tpc.F_FindWindowEx(iIE_Handl, 0, "Frame Tab", "");
        int iIE_Handl_3 = tpc.F_FindWindowEx(iIE_Handl_2, 0, "TabWindowClass", "Node : " + dev.ModuleType + "-AUTOTEST - Internet Explorer");    // 注意是CTestSCADA而不是TestSCADA  Jammy這邊要改!!
        int iIE_Handl_4 = tpc.F_FindWindowEx(iIE_Handl_3, 0, "Shell DocObject View", "");
        int iIE_Handl_5 = tpc.F_FindWindowEx(iIE_Handl_4, 0, "Internet Explorer_Server", "");
        int iIE_Handl_6 = tpc.F_FindWindowEx(iIE_Handl_5, 0, "AfxOleControl42s", "");
        int iIE_Handl_7 = tpc.F_FindWindowEx(iIE_Handl_6, 0, "AfxWnd42s", "");
        int iWA_MainPage = tpc.F_FindWindowEx(iIE_Handl_7, 0, "ActXBroadWinBwviewWClass", "Advantech View 001 - main:untitled");

        if (iWA_MainPage > 0)
        {
            System.Threading.Thread.Sleep(5000);
            tpc.F_PostMessage(iWA_MainPage, tpc.V_WM_KEYDOWN, tpc.V_VK_ESCAPE, 0);
            System.Threading.Thread.Sleep(1000);
        }

        // Login keyboard
        PrintTitle("Login keyboard");
        int iLoginKeyboard_Handle = tpc.F_FindWindow("#32770", "Login");
        int iEnterText = tpc.F_FindWindowEx(iLoginKeyboard_Handle, 0, "Edit", "");
        if (iEnterText > 0)
        {
            SendCharToHandle(iEnterText, 100, "admin");
            tpc.F_PostMessage(iEnterText, tpc.V_WM_KEYDOWN, tpc.V_VK_RETURN, 0);
        }

        System.Threading.Thread.Sleep(1000);
        SendKeys.SendWait("^{F5}");
        System.Threading.Thread.Sleep(1000);

        int iPointInfo_Handle = tpc.F_FindWindow("#32770", "Point Info");
        int iEnterText_PointInfo = tpc.F_FindWindowEx(iPointInfo_Handle, 0, "Edit", "");
        if (iEnterText_PointInfo > 0)
        {
            if (ai_point > 0)
            {
                // AT_AI0
                PrintTitle("Print Screen AI Tag.");
                SendCharToHandle(iEnterText_PointInfo, 100, "AUTOTEST-AI_0");
                System.Threading.Thread.Sleep(1500);

                PrintScreen("PlugandPlay_TagInfoSyncTest_AUTOTEST-AI_0", System.IO.Directory.GetCurrentDirectory());
                for (int i = 1; i <= 20; i++)
                {
                    tpc.F_PostMessage(iEnterText_PointInfo, tpc.V_WM_KEYDOWN, tpc.V_VK_BACK, 0);
                    System.Threading.Thread.Sleep(100);
                }
            }

            if (di_point > 0)
            {
                PrintTitle("Print Screen DI Tag.");
                SendCharToHandle(iEnterText_PointInfo, 100, "AUTOTEST-DI_0");
                System.Threading.Thread.Sleep(1500);

                PrintScreen("PlugandPlay_TagInfoSyncTest_AUTOTEST-DI_0", System.IO.Directory.GetCurrentDirectory());
                for (int i = 1; i <= 20; i++)
                {
                    tpc.F_PostMessage(iEnterText_PointInfo, tpc.V_WM_KEYDOWN, tpc.V_VK_BACK, 0);
                    System.Threading.Thread.Sleep(100);
                }
            }

            if (do_point > 0)
            {
                // AT_AI0
                PrintTitle("Print Screen DO Tag.");
                SendCharToHandle(iEnterText_PointInfo, 100, "AUTOTEST-DO_0");
                System.Threading.Thread.Sleep(1500);

                PrintScreen("PlugandPlay_TagInfoSyncTest_AUTOTEST-DO_0", System.IO.Directory.GetCurrentDirectory());
                for (int i = 1; i <= 20; i++)
                {
                    tpc.F_PostMessage(iEnterText_PointInfo, tpc.V_WM_KEYDOWN, tpc.V_VK_BACK, 0);
                    System.Threading.Thread.Sleep(100);
                }
            }
            
            
            // AT_DO0100
            //PrintTitle("Exe4012E_AISetting");
            //SendCharToHandle(iEnterText_PointInfo, 100, "AUTOTEST-DO0");
            //System.Threading.Thread.Sleep(1500);

            //PrintScreen("PlugandPlay_TagInfoSyncTest_AUTOTEST-DO0", System.IO.Directory.GetCurrentDirectory());
            //for (int i = 1; i <= 20; i++)
            //{
            //    tpc.F_PostMessage(iEnterText_PointInfo, tpc.V_WM_KEYDOWN, tpc.V_VK_BACK, 0);
            //    System.Threading.Thread.Sleep(100);
            //}

        }

        tpc.F_PostMessage(iPointInfo_Handle, tpc.V_WM_KEYDOWN, tpc.V_VK_ESCAPE, 0);
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

    private void PrintScreen(string sFileName, string sFilePath)
    {
        PrintTitle("PrintScreen");
        Bitmap myImage = new Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height);
        Graphics g = Graphics.FromImage(myImage);
        g.CopyFromScreen(new Point(0, 0), new Point(0, 0), new Size(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height));
        IntPtr dc1 = g.GetHdc();
        g.ReleaseHdc(dc1);
        //myImage.Save(@"c:\screen0.jpg");
        myImage.Save(string.Format("{0}\\{1}_{2:yyyyMMdd_hhmmss}.jpg", sFilePath, sFileName, DateTime.Now));
    }

    private void StartNode()
    {
        PrintTitle("StartNode");
        api.SwitchToCurWindow(0);
        api.SwitchToFrame("rightFrame", 0);
        api.ByXpath("//tr[2]/td/a[5]/font").Click();    // start kernel

        string main; object subobj;
        api.GetWinHandle(out main, out subobj);
        IEnumerator<String> windowIterator = (IEnumerator<String>)subobj;

        List<string> items = new List<string>();
        while (windowIterator.MoveNext())
            items.Add(windowIterator.Current);

        if (main != items[1])
        {
            api.SwitchToWinHandle(items[1]);
        }
        else
        {
            api.SwitchToWinHandle(items[0]);
        }
        api.ByName("submit").Enter("").Submit().Exe();
        PrintStep();
        System.Threading.Thread.Sleep(5000);    // Wait 30s for start kernel finish
        api.Close();
        api.SwitchToWinHandle(main);        // switch back to original window

        PrintStep();
    }






}
