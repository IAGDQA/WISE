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
    int errorCnt = 0;
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
        if (!AIModChkBlock())
        {
            PrintTitle("Module not support.");
            return;
        }
        //
        errorCnt = 0;
        selenium = new AdvSeleniumAPIv2();

        selenium.StartupServer("http://" + textBox1.Text);
        System.Threading.Thread.Sleep(1000);

        selenium.Type("id=ACT0", "root");
        selenium.Type("id=PWD0", "00000000");
        selenium.Click("id=APY0");
        selenium.WaitForPageToLoad("30000");
        PrintStep();
        //
        UIConfgBlock();
        //AIConfgStsBlock();
        AIConfgChnBlock();
        AIConfgComBlock();
        
        selenium.Close();

        
    }
    int ai_num = 0;
    bool AIModChkBlock()
    {
        if (dev.ModuleType.ToUpper() == "WISE-4012E")
        {
            ai_num = 2; return true;
        }
        else if (dev.ModuleType.ToUpper() == "WISE-4012"
                  || dev.ModuleType.ToUpper() == "WISE-4010/LAN")
        {
            ai_num = 4; return true;
        }
        return false;
    }

    void UIConfgBlock()
    {
        if (dev.ModuleType.ToUpper() == "WISE-4012")
        {
            PrintTitle("UIConfg");
            selenium.Click("id=ioStatus0");
            for (int i = 0; i < ai_num; i++)
            {
                selenium.Select("id=selMd_" + i.ToString(), "label=AI");
            }
            selenium.Click("id=btnUIConfig");
            PrintStep();
        }
    }

    void AIConfgStsBlock()
    {
        //PrintTitle("AIConfgSts");
        //api.ById("ioStatus0").ClickAndWait(1000);
        //if (dev.ModuleType.ToUpper() == "WISE-4012") api.ByTxt("AI").ClickAndWait(1000);
        ////
        //api.ById("btnLoA").Click();
        //api.ById("btnHiA").Click();
        //api.Wait(1000);
        ////
        //for (int i = 1; i <= ai_num; i++)
        //{
        //    api.Select(i).ByXpath("//select[@id='selCh']").Exe();
        //    api.Wait(1000);
        //    api.ById("btnLoA").Click();
        //    api.ById("btnHiA").Click();
        //    api.Wait(1000);
        //}
        //PrintStep();
        ////
        //if (dev.ModuleType.ToUpper() == "WISE-4012") api.ByTxt("AI").ClickAndWait(1000);
        //api.ByXpath("//a[contains(text(),'Max')]").ClickAndWait(1000);
        //for (int i = 1; i <= ai_num; i++)
        //{
        //    api.Select(i).ByXpath("(//select[@id='selCh'])[2]").Exe();
        //    api.Wait(1000);
        //    api.ByCss("#AIMaxStatusTable > div.panel-body > div.table-responsive > center "
        //    + "> table.table-control > tbody > tr.aiMinMaxType > td.td-control > #btnClrL").ClickAndWait(1000);
        //}
        //PrintStep();
        ////
        //if (dev.ModuleType.ToUpper() == "WISE-4012") api.ByTxt("AI").ClickAndWait(1000);
        //api.ByXpath("//a[contains(text(),'Min')]").ClickAndWait(1000);
        //for (int i = 1; i <= ai_num; i++)
        //{
        //    api.Select(i).ByXpath("(//select[@id='selCh'])[3]").Exe();
        //    api.Wait(1000);
        //    api.ByCss("#AIMinStatusTable > div.panel-body > div.table-responsive > center "
        //    + "> table.table-control > tbody > tr.aiMinMaxType > td.td-control > #btnClrL").ClickAndWait(1000);
        //}
        //PrintStep();
        //
        //api.Select(1).ByXpath("//select[@id='selCh']").Exe();
        //api.Wait(1000);
        //api.Select(2).ByXpath("//select[@id='selCh']").Exe();
        //api.Wait(1000);
        //api.Select(3).ByXpath("//select[@id='selCh']").Exe();
        //api.Wait(1000);
        //api.Select(4).ByXpath("//select[@id='selCh']").Exe();
        //api.Wait(1000);
        //api.ById("btnLoA").Click();
        //api.ById("btnHiA").Click();
        //api.Wait(1000);
        //
        //api.ByTxt("AI").ClickAndWait(1000);
        //api.ByXpath("//a[contains(text(),'Max')]").ClickAndWait(1000);
        //api.Select(1).ByXpath("(//select[@id='selCh'])[2]").Exe();
        //api.Wait(1000);
        //api.Select(2).ByXpath("(//select[@id='selCh'])[2]").Exe();
        //api.Wait(1000);
        //api.Select(3).ByXpath("(//select[@id='selCh'])[2]").Exe();
        //api.Wait(1000);
        //api.Select(4).ByXpath("(//select[@id='selCh'])[2]").Exe();
        //api.Wait(1000);
        //api.ByCss("#AIMaxStatusTable > div.panel-body > div.table-responsive > center "
        //    +"> table.table-control > tbody > tr.aiMinMaxType > td.td-control > #btnClrL").ClickAndWait(1000);
        //
        //api.ByTxt("AI").ClickAndWait(1000);
        //api.ByXpath("//a[contains(text(),'Min')]").ClickAndWait(1000);
        //api.Select(1).ByXpath("(//select[@id='selCh'])[3]").Exe();
        //api.Wait(1000);
        //api.Select(2).ByXpath("(//select[@id='selCh'])[3]").Exe();
        //api.Wait(1000);
        //api.Select(3).ByXpath("(//select[@id='selCh'])[3]").Exe();
        //api.Wait(1000);
        //api.Select(4).ByXpath("(//select[@id='selCh'])[3]").Exe();
        //api.Wait(1000);
        //api.ByCss("#AIMinStatusTable > div.panel-body > div.table-responsive > center "
        //    +"> table.table-control > tbody > tr.aiMinMaxType > td.td-control > #btnClrL").ClickAndWait(1000);
        //
    }

    void AIConfgChnBlock()
    {
        PrintTitle("AIConfgChn");
        selenium.Click("id=ioStatus0");
        selenium.Click("link=AI");
        selenium.Click("//a[contains(text(),'Configuration')]");

        //
        if (dev.ModuleType.ToUpper() == "WISE-4012")
        {
            Exe4012_AISetting();
        }
        else if (dev.ModuleType.ToUpper() == "WISE-4012E")
        {
            Exe4012E_AISetting();
        }
        else if (dev.ModuleType.ToUpper() == "WISE-4010/LAN")
        {
            Exe4010LAN_AISetting();
        }

        PrintStep();
    }

    void AIConfgComBlock()
    {
        if (dev.ModuleType.ToUpper() == "WISE-4012"
            || dev.ModuleType.ToUpper() == "WISE-4012E")
        {
            WISE_AIConfgComBlock();
        }
        else if (dev.ModuleType.ToUpper() == "WISE-4010/LAN")
        {
            WISE4010LAN_AIConfgComBlock();
        }
    }
    void WISE_AIConfgComBlock()
    {
        PrintTitle("AIConfgComBlock");
        selenium.Click("id=ioStatus0");
        selenium.Click("link=AI");
        selenium.Click("//a[contains(text(),'Configuration')]");
        selenium.Click("link=Common Settings");
        selenium.Click("id=inpEnB");
        selenium.Select("id=selBMd", "label=Down scale");
        selenium.Click("xpath=(//input[@type='checkbox'])[13]");
        selenium.Click("xpath=(//input[@type='checkbox'])[14]");
        selenium.Click("xpath=(//input[@type='checkbox'])[15]");
        selenium.Click("xpath=(//input[@type='checkbox'])[16]");
        selenium.Click("id=btnSubmit");
        selenium.Click("css=input.avgMType");
        selenium.Click("css=input.avgMType");
        selenium.Click("id=btnSubmit");

        PrintStep();
    }
    void WISE4010LAN_AIConfgComBlock()
    {
        PrintTitle("WISE4010LAN_AIConfgComBlock");
        //selenium.Click("id=ioStatus0");
        //selenium.Click("link=AI");
        //selenium.Click("//a[contains(text(),'Configuration')]");
        selenium.Click("link=Common Settings");
        selenium.Click("id=inpEnB");
        selenium.Select("id=selBMd", "label=Down scale");
        selenium.Select("id=selSmp", "label=1000 Hz/Ch");
        selenium.Click("xpath=(//input[@type='checkbox'])[9]");
        selenium.Click("xpath=(//input[@type='checkbox'])[10]");
        selenium.Click("xpath=(//input[@type='checkbox'])[11]");
        selenium.Click("xpath=(//input[@type='checkbox'])[12]");
        selenium.Click("id=btnSubmit");

        PrintStep();
    }

    void Exe4012_AISetting()
    {
        PrintTitle("Exe4012_AISetting");
        for (int i = 0; i < ai_num; i++)
        {
            selenium.Select("css=div.input-group > #selCh", "label=" + i.ToString());
            selenium.Type("id=inpTag", "ABCDEFGHIJKLMNOPQRST" + i.ToString());
            selenium.Select("id=selRng", "label=+/- 150 mV");
            selenium.Click("id=inpEn");
            selenium.Click("id=inpEn");
            selenium.Type("id=inpLoS", "999" + i.ToString());
            selenium.Type("id=inpHiS", "9999" + i.ToString());
            selenium.Type("id=inpLoP", "888" + i.ToString());
            selenium.Type("id=inpHiP", "8888" + i.ToString());
            selenium.Type("id=inpUni", "1000");
            selenium.Click("id=inpEnLA");
            selenium.Select("id=selLAMd", "label=Latch");
            selenium.Type("css=div.input-group > #inpLoA", "1");
            selenium.Click("id=inpEnHA");
            selenium.Select("id=selHAMd", "label=Latch");
            selenium.Type("css=div.input-group > #inpHiA", "2");
            selenium.Click("css=#aiConfigBaseForm > div.panel-footer.clearfix > div.pull-right > #btnSubmit");
            PrintStep();
        }
        //average
        selenium.Select("css=div.input-group > #selCh", "label=Average");
        selenium.Type("id=inpLoS", "999");
        selenium.Type("id=inpHiS", "9999");
        selenium.Type("id=inpLoP", "888");
        selenium.Type("id=inpHiP", "8888");
        selenium.Type("id=inpUni", "1000");
        selenium.Click("id=inpEnLA");
        selenium.Select("id=selLAMd", "label=Latch");
        selenium.Type("css=div.input-group > #inpLoA", "1");
        selenium.Click("id=inpEnHA");
        selenium.Select("id=selHAMd", "label=Latch");
        selenium.Type("css=div.input-group > #inpHiA", "2");
        selenium.Click("css=#aiConfigBaseForm > div.panel-footer.clearfix > div.pull-right > #btnSubmit");
        PrintStep();
        //all
        selenium.Select("css=div.input-group > #selCh", "label=All");
        selenium.Type("id=inpLoS", "666");
        selenium.Type("id=inpHiS", "6666");
        selenium.Type("id=inpLoP", "777");
        selenium.Type("id=inpHiP", "7777");
        selenium.Type("id=inpUni", "1000");
        selenium.Click("id=inpEnLA");
        selenium.Select("id=selLAMd", "label=Latch");
        selenium.Type("css=div.input-group > #inpLoA", "1");
        selenium.Click("id=inpEnHA");
        selenium.Select("id=selHAMd", "label=Latch");
        selenium.Type("css=div.input-group > #inpHiA", "2");
        selenium.Click("css=#aiConfigBaseForm > div.panel-footer.clearfix > div.pull-right > #btnSubmit");
        PrintStep();
    }

    void Exe4012E_AISetting()
    {
        PrintTitle("Exe4012E_AISetting");


    }
    void Exe4010LAN_AISetting()
    {
        PrintTitle("Exe4010LAN_AISetting");
        for (int i = 0; i < ai_num; i++)
        {
            selenium.Select("css=div.input-group > #selCh", "label=" + i.ToString());
            selenium.Type("id=inpTag", "ABCDEFGHIJKLMNOPQRST" + i.ToString());
            selenium.Select("id=selRng", "label=0 ~ 20 mA");
            selenium.Click("id=inpEn");
            selenium.Click("id=inpEn");
            selenium.Type("id=inpLoS", "999" + i.ToString());
            selenium.Type("id=inpHiS", "9999" + i.ToString());
            selenium.Type("id=inpLoP", "888" + i.ToString());
            selenium.Type("id=inpHiP", "8888" + i.ToString());
            selenium.Type("id=inpUni", "1000");
            selenium.Click("id=inpEnLA");
            selenium.Select("id=selLAMd", "label=Latch");
            selenium.Type("css=div.input-group > #inpLoA", "1");
            selenium.Click("id=inpEnHA");
            selenium.Select("id=selHAMd", "label=Latch");
            selenium.Type("css=div.input-group > #inpHiA", "2");
            selenium.Click("css=#aiConfigBaseForm > div.panel-footer.clearfix > div.pull-right > #btnSubmit");
            PrintStep();
        }
        //average
        selenium.Select("css=div.input-group > #selCh", "label=Average");
        selenium.Type("id=inpTag", "ABCDEFGHIJKLMNOPQRSTU");
        selenium.Type("id=inpLoS", "999");
        selenium.Type("id=inpHiS", "9999");
        selenium.Type("id=inpLoP", "888");
        selenium.Type("id=inpHiP", "8888");
        selenium.Type("id=inpUni", "1000");
        selenium.Click("id=inpEnLA");
        selenium.Select("id=selLAMd", "label=Latch");
        selenium.Type("css=div.input-group > #inpLoA", "1");
        selenium.Click("id=inpEnHA");
        selenium.Select("id=selHAMd", "label=Latch");
        selenium.Type("css=div.input-group > #inpHiA", "2");
        selenium.Click("css=#aiConfigBaseForm > div.panel-footer.clearfix > div.pull-right > #btnSubmit");
        PrintStep();
        //all
        selenium.Select("css=div.input-group > #selCh", "label=All");
        selenium.Type("id=inpLoS", "666");
        selenium.Type("id=inpHiS", "6666");
        selenium.Type("id=inpLoP", "777");
        selenium.Type("id=inpHiP", "7777");
        selenium.Type("id=inpUni", "1000");
        selenium.Click("id=inpEnLA");
        selenium.Select("id=selLAMd", "label=Momentary");
        selenium.Type("css=div.input-group > #inpLoA", "10");
        selenium.Click("id=inpEnHA");
        selenium.Select("id=selHAMd", "label=Momentary");
        selenium.Type("css=div.input-group > #inpHiA", "20");
        selenium.Click("css=#aiConfigBaseForm > div.panel-footer.clearfix > div.pull-right > #btnSubmit");
        PrintStep();
    }

}//class
