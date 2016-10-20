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
            En = 1,
            Fmt = 0,
            Fn = 1,
            TmF = 1,
            DEn = 2,
            DTim = 864000,
            DItm = 10000,
            SEn = 2,
            STim = 864000,
            SItm = 1000,
            DTag = "123456789012345678901234567890123456789012345678901234567890"
                    + "123456789012345678901234567890123456789012345678901234567890"
                    + "ABCDEFGH",
            STag = "123456789012345678901234567890123456789012345678901234567890"
                    + "123456789012345678901234567890123456789012345678901234567890"
                    + "ABCDEFGH",
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
                                        "http://" + Device.IPAddress + "/log_upload");
        //
        if (changeFlg) servAct = ServiceAction.GetNetConfig_ag;
        //
        ExeRes = new wResult()
        {
            Method = HttpRequestOption.GET,
            Ins = WISE_RESTFUL_URI.log_upload,
        };
        
    }

    private void PatchSysInfoRequest()//Patch info
    {
        Print(new wResult() { Des = "PatchSysInfoRequest" });
        servAct = ServiceAction.PatchSysInfo;

        JavaScriptSerializer serializer = new JavaScriptSerializer();
        string sz_Jsonify = serializer.Serialize(ChangeDataArry);

        m_HttpRequest.SendPATCHRequest(Device.Account, Device.Password, GetURL(Device.IPAddress, Device.Port
                                    , WISE_RESTFUL_URI.log_upload.ToString()), sz_Jsonify);
        changeFlg = true;
        //
        ExeRes = new wResult()
        {
            Method = HttpRequestOption.PATCH,
            Ins = WISE_RESTFUL_URI.log_upload,
            Res = ExeCaseRes.Pass,
        }; Print(ExeRes);

        this.InvokeWaitStep();
    }

    private void VerifyItems()
    {
        int errorCnt = 0; changeFlg = false;
        Print(new wResult() { Des = "VerifyItems" });
        bool chk = false;
        if (GetDataArry.En != ChangeDataArry.En) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "En  check [" + GetDataArry.En + "]", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        chk = false;
        if (GetDataArry.Fn != ChangeDataArry.Fn) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "Fn   check [" + GetDataArry.Fn + "]", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        chk = false;
        if (GetDataArry.TmF != ChangeDataArry.TmF) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "TmF   check [" + GetDataArry.TmF + "]", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        chk = false;
        if (GetDataArry.DEn != ChangeDataArry.DEn) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "DEn  check [" + GetDataArry.DEn + "]", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        chk = false;
        if (GetDataArry.DTim != ChangeDataArry.DTim) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "DTim  check [" + GetDataArry.DTim + "]", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        chk = false;
        if (GetDataArry.DItm != ChangeDataArry.DItm) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "DItm  check [" + GetDataArry.DItm + "]", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        chk = false;
        if (GetDataArry.SEn != ChangeDataArry.SEn) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "SEn  check [" + GetDataArry.SEn + "]", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        chk = false;
        if (GetDataArry.STim != ChangeDataArry.STim) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "STim  check [" + GetDataArry.STim + "]", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        chk = false;
        if (GetDataArry.SItm != ChangeDataArry.SItm) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "SItm   check [" + GetDataArry.SItm + "]", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        chk = false;
        if (GetDataArry.DTag != ChangeDataArry.DTag) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "DTag   check [" + GetDataArry.DTag + "]", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        chk = false;
        if (GetDataArry.STag != ChangeDataArry.STag) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "STag   check [" + GetDataArry.STag + "]", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });

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
            GetDataArry.En = data.En;
            GetDataArry.Fmt = data.Fmt;
            GetDataArry.Fn = data.Fn;
            GetDataArry.TmF = data.TmF;
            GetDataArry.DEn = data.DEn;
            GetDataArry.DTim = data.DTim;
            GetDataArry.DItm = data.DItm;
            GetDataArry.SEn = data.SEn;
            GetDataArry.STim = data.STim;
            GetDataArry.SItm = data.SItm;
            GetDataArry.DTag = data.DTag;
            GetDataArry.STag = data.STag;
            
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
        public int En { get; set; }
        public int Fmt { get; set; }
        public int Fn { get; set; }
        public int TmF { get; set; }
        public int DEn { get; set; }
        public int DTim { get; set; }
        public int DItm { get; set; }
        public int SEn { get; set; }
        public int STim { get; set; }
        public int SItm { get; set; }
        public string DTag { get; set; }
        public string STag { get; set; }
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

