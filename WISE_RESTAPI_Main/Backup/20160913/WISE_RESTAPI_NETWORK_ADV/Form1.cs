using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

using System.Net;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Threading;
using System.Web.Script.Serialization;
using Model;
using Service;

public partial class Form1 : Form
{
    DeviceModel Device;
    private AdvantechHttpWebUtility m_HttpRequest;
    DataHandleService dataHld;
    bool DevConnFail = false;
    ServiceAction servAct = new ServiceAction();

    private delegate void DataGridViewCtrlAddDataRow(DataGridViewRow i_Row);
    private DataGridViewCtrlAddDataRow m_DataGridViewCtrlAddDataRow;
    internal const int Max_Rows_Val = 65535;

    SysData GetDataArry = new SysData();
    SysData ChangeDataArry = new SysData();//change description content
    bool changeFlg = false;
    wResult ExeRes;

    iALibrary.iAS ias = new iALibrary.iAS();


    public Form1()
    {
        InitializeComponent();

        ChangeDataArry = new SysData()
        {
            Idl = 999,
            CWDT = 1,
            PStr = 7968,
            PpG = 6689,
            PWeb = 800,
            EnSNTP = 1,
            SNTP1 = "123456789012345678901234567890123456789012345678901234567890"
                                + "ABC",
            SNTP2 = "123456789012345678901234567890123456789012345678901234567890"
                                + "ABC",
            SNTPI = 1152,
            TZHr = 6,
            TZMn = 30,
            TZC = 49,
        };
    }

    private void Form1_Load(object sender, EventArgs e)
    {
        //Get IO value
        m_HttpRequest = new AdvantechHttpWebUtility();
        m_HttpRequest.ResquestOccurredError += this.OnGetHttpRequestError;
        m_HttpRequest.ResquestResponded += this.OnGetData;
        //
        Device = new DeviceModel()//20150626 建立一個DeviceModel給所有Service
        {
            IPAddress = textBox1.Text,
            Account = "root",
            Password = "00000000",
            Port = 80,
            SlotNum = 0,
            ModbusAddr = 1,
            ModbusTimeOut = 3000,
        };
        //
        dataGridView1.ColumnHeadersVisible = true;
        DataGridViewTextBoxColumn newCol = new DataGridViewTextBoxColumn(); // add a column to the grid
        newCol.HeaderText = "Time";
        newCol.Name = "clmTs";
        newCol.Visible = true;
        newCol.Width = 50;
        dataGridView1.Columns.Add(newCol);
        //
        newCol = new DataGridViewTextBoxColumn();
        newCol.HeaderText = "Method";
        newCol.Name = "clmStp";
        newCol.Visible = true;
        newCol.Width = 50;
        dataGridView1.Columns.Add(newCol);
        //
        newCol = new DataGridViewTextBoxColumn();
        newCol.HeaderText = "Instruction";
        newCol.Name = "clmIns";
        newCol.Visible = true;
        newCol.Width = 100;
        dataGridView1.Columns.Add(newCol);        
        //
        newCol = new DataGridViewTextBoxColumn();
        newCol.HeaderText = "Description";
        newCol.Name = "clmDes";
        newCol.Visible = true;
        newCol.Width = 100;
        dataGridView1.Columns.Add(newCol);
        //
        newCol = new DataGridViewTextBoxColumn();
        newCol.HeaderText = "Result";
        newCol.Name = "clmRes";
        newCol.Visible = true;
        newCol.Width = 80;
        dataGridView1.Columns.Add(newCol);
        //
        newCol = new DataGridViewTextBoxColumn();
        newCol.HeaderText = "Error";
        newCol.Name = "clmErr";
        newCol.Visible = true;
        newCol.Width = 100;
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

        dataHld = new DataHandleService();
        textBox1.Text = dataHld.GetPara(Application.StartupPath);
        //debug
        //GetNetConfigRequest();
    }

    protected override void WndProc(ref System.Windows.Forms.Message iMessage)
    {
        int iType = (int)iMessage.WParam;

        if (iMessage.Msg == ias.MsgExcutionCtrlCommand)
        {
            switch (iType)
            {
                case (int)iALibrary.iCommand.Start:
                    RunProc();
                    break;
                case (int)iALibrary.iCommand.Stop:

                    ias.iStatus(iALibrary.iStatus.Stop);
                    break;
            }
        }
        base.WndProc(ref iMessage);
    }

    private void RunProc()
    {
        ias.iStatus(iALibrary.iStatus.Running);
        textBox1.Text = dataHld.GetPara(Application.StartupPath);
        GetNetConfigRequest();
    }

    private string GetURL(string ip, int port, string requestUri)
    {
        return "http://" + ip + ":" + port.ToString() + "/" + requestUri;
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

    private void OnGetHttpRequestError(Exception ex)
    {
        ExeRes.Res = ExeCaseRes.Fail;
        ExeRes.Err = ex.ToString();
        Print(ExeRes);
    }

    private void OnGetData(string rawData)//Feedback Http request
    {
        switch (servAct)
        {
            case ServiceAction.GetNetConfig:
                var Obj01 = AdvantechHttpWebUtility.ParserJsonToObj<SysData>(rawData);
                UpdateDevUIStatus(Obj01);
                //
                ExeRes.Res = ExeCaseRes.Pass;Print(ExeRes);
                this.InvokeWaitStep();
                break;
            case ServiceAction.PatchSysInfo:
                break;
            case ServiceAction.GetNetConfig_ag:
                var Obj03 = AdvantechHttpWebUtility.ParserJsonToObj<SysData>(rawData);
                UpdateDevUIStatus(Obj03);
                //
                ExeRes.Res = ExeCaseRes.Pass; Print(ExeRes);
                this.InvokeWaitStep();
                break;
        }
    }

    void Print(wResult obj)
    {
        DataGridViewRow dgvRow;
        DataGridViewCell dgvCell;
        dgvRow = new DataGridViewRow();
        //dgvRow.DefaultCellStyle.Font = new Font(this.Font, FontStyle.Regular);
        dgvCell = new DataGridViewTextBoxCell(); //Column Time
        var dataTimeInfo = DateTime.Now.ToString("yyyy-MM-dd HH:MM:ss");
        dgvCell.Value = dataTimeInfo;
        dgvRow.Cells.Add(dgvCell);
        //
        dgvCell = new DataGridViewTextBoxCell();
        dgvCell.Value = obj.Method;
        dgvRow.Cells.Add(dgvCell);
        //
        dgvCell = new DataGridViewTextBoxCell();
        dgvCell.Value = obj.Ins;
        dgvRow.Cells.Add(dgvCell);
        //
        dgvCell = new DataGridViewTextBoxCell();
        dgvCell.Value = obj.Des;
        dgvRow.Cells.Add(dgvCell);
        //
        dgvCell = new DataGridViewTextBoxCell();
        dgvCell.Value = obj.Res;
        dgvRow.Cells.Add(dgvCell);
        //
        dgvCell = new DataGridViewTextBoxCell();
        dgvCell.Value = obj.Err;
        dgvRow.Cells.Add(dgvCell);

        m_DataGridViewCtrlAddDataRow(dgvRow);

        ExeRes = new wResult();
    }

    private void button1_Click(object sender, EventArgs e)
    {
        GetNetConfigRequest();
    }

    private void InvokeWaitStep()//start to read IO invoke
    {
        int m_iPollingTime = 500;
        Thread.Sleep(m_iPollingTime);//m_iPollingTime
        NextStep();
    }

    private void NextStep()
    {
        if (servAct == ServiceAction.GetNetConfig_ag)
        {
            VerifyItems();
        }
        else if (servAct == ServiceAction.PatchSysInfo)
        {
            GetNetConfigRequest();
        }
        else if (servAct == ServiceAction.GetNetConfig)
        {
            PatchSysInfoRequest();
        }
    }

    //Request Cmd
    private void GetNetConfigRequest()
    {
        Print(new wResult() { Des = "GetNetConfigRequest" });
        dataHld.SavePara(Application.StartupPath, textBox1.Text);
        Device.IPAddress = textBox1.Text;
        servAct = ServiceAction.GetNetConfig;
        m_HttpRequest.SendGETRequest(Device.Account, Device.Password,
                                        GetURL(Device.IPAddress, Device.Port
                                    , WISE_RESTFUL_URI.net_config.ToString()));
        //
        if (changeFlg) servAct = ServiceAction.GetNetConfig_ag;
        //
        ExeRes = new wResult()
        {
            Method = HttpRequestOption.GET,
            Ins = WISE_RESTFUL_URI.net_config,
        };
        
    }

    private void PatchSysInfoRequest()//Patch info
    {
        Print(new wResult() { Des = "PatchSysInfoRequest" });
        servAct = ServiceAction.PatchSysInfo;

        JavaScriptSerializer serializer = new JavaScriptSerializer();
        string sz_Jsonify = serializer.Serialize(ChangeDataArry);

        m_HttpRequest.SendPATCHRequest(Device.Account, Device.Password, GetURL(Device.IPAddress, Device.Port
                                    , WISE_RESTFUL_URI.net_config.ToString()), sz_Jsonify);
        changeFlg = true;
        //
        ExeRes = new wResult()
        {
            Method = HttpRequestOption.PATCH,
            Ins = WISE_RESTFUL_URI.net_config,
            Res = ExeCaseRes.Pass,
        }; Print(ExeRes);

        this.InvokeWaitStep();
    }

    private void VerifyItems()
    {
        int errorCnt = 0; changeFlg = false;
        Print(new wResult() { Des = "VerifyItems" });
        bool chk = false;
        if (GetDataArry.Idl != ChangeDataArry.Idl) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "Idl check", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        chk = false;
        if (GetDataArry.CWDT != ChangeDataArry.CWDT) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "CWDT check", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        //chk = false;
        //if (GetDataArry.PStr != ChangeDataArry.PStr) { chk = true; errorCnt++; }
        //Print(new wResult() { Des = "PStr check", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        //chk = false;
        //if (GetDataArry.PpG != ChangeDataArry.PpG) { chk = true; errorCnt++; }
        //Print(new wResult() { Des = "PpG check", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        chk = false;
        if (GetDataArry.PWeb != ChangeDataArry.PWeb) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "PWeb check", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        chk = false;
        if (GetDataArry.EnSNTP != ChangeDataArry.EnSNTP) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "EnSNTP check", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        chk = false;
        if (GetDataArry.SNTP1 != ChangeDataArry.SNTP1) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "SNTP1 check", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        chk = false;
        if (GetDataArry.SNTP2 != ChangeDataArry.SNTP2) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "SNTP2 check", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        chk = false;
        if (GetDataArry.SNTPI != ChangeDataArry.SNTPI) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "SNTPI check", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        chk = false;
        if (GetDataArry.TZC != ChangeDataArry.TZC) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "TZC check", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        chk = false;
        if (GetDataArry.TZHr != ChangeDataArry.TZHr) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "TZHr check", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        chk = false;
        if (GetDataArry.TZMn != ChangeDataArry.TZMn) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "TZMn check", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });

        //Return the test result
        if (errorCnt > 0) ias.ReturnRes((int)iALibrary.iResult.Fail);
        else ias.ReturnRes((int)iALibrary.iResult.Pass);
        //To notify iATester the test is completion
        ias.iStatus(iALibrary.iStatus.Completion);
    }

    #region ---- Update UI ----
    private void UpdateDevUIStatus(SysData data)
    {
        try
        {
            GetDataArry.Idl = data.Idl;
            GetDataArry.CWDT = data.CWDT;
            GetDataArry.PStr = data.PStr;
            GetDataArry.PpG = data.PpG;
            GetDataArry.PWeb = data.PWeb;
            GetDataArry.EnSNTP = data.EnSNTP;
            GetDataArry.SNTP1 = data.SNTP1;
            GetDataArry.SNTP2 = data.SNTP2;
            GetDataArry.SNTPI = data.SNTPI;
            GetDataArry.TZHr = data.TZHr;
            GetDataArry.TZMn = data.TZMn;
            GetDataArry.TZC = data.TZC;
        }
        catch (Exception e)
        {
            //OnGetDevHttpRequestError(e);
        }
        finally
        {            
            System.GC.Collect();
        }
    }
    #endregion

    public class SysData
    {
        public int Idl { get; set; }
        public int CWDT { get; set; }
        public int PStr { get; set; }
        public int PpG { get; set; }
        public int PWeb { get; set; }
        public int EnSNTP { get; set; }
        public string SNTP1 { get; set; }
        public string SNTP2 { get; set; }
        public int SNTPI { get; set; }
        public double TZHr { get; set; }
        public double TZMn { get; set; }
        public double TZC { get; set; }
    }   

    public enum ServiceAction
    {
        Idel = 0,
        GetNetConfig = 1,
        PatchSysInfo = 2,
        GetNetConfig_ag = 3,
        Verify = 4,

        Done = 99,
    }
}

