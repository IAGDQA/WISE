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
using iATester;

public partial class Form1 : Form, iATester.iCom
{
    DeviceModel Device;
    private AdvantechHttpWebUtility m_HttpRequest;
    DataHandleService dataHld;
    ServiceAction servAct = new ServiceAction();

    private delegate void DataGridViewCtrlAddDataRow(DataGridViewRow i_Row);
    private DataGridViewCtrlAddDataRow m_DataGridViewCtrlAddDataRow;
    internal const int Max_Rows_Val = 65535;

    SysCoilData GetDataArry = new SysCoilData();
    SysCoilData ChangeDataArry = new SysCoilData();//change description content
    SysRegData GetRegDataArry = new SysRegData();
    SysRegData ChangeRegDataArry = new SysRegData();//change description content
    bool changeFlg = false;
    wResult ExeRes;
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

        ChangeDataArry = new SysCoilData()
        {
            Md=0,
            DI = 11,
            CtS = 331,
            CtClr = 371,
            CtOv = 411,
            Lch = 451,
            DO = 171,
            AIHR = 1011,
            AILR = 1111,
            AIB = 1211,
            HAlm = 1311,
            LAlm = 1411,
            GCtClr = 351,
            LB = 9999,
            ExB = 1001,
        };
        ChangeRegDataArry = new SysRegData()
        {
            Md = 0,
            DI = 3011,
            CtFq = 100,
            DO = 2301,
            PsLo = 2001,
            PsHi = 2401,
            PsAV = 2601,
            PsIV = 2801,
            AI = 3001,
            HisH = 3101,
            HisL = 3501,
            AIF = 3601,
            HisHF = 3701,
            HisLF = 3801,
            AIFl = 4101,
            AICd = 4201,
            AICh = 3301,
            AISc = 3901,
            AIPF = 1231,
            AO = 7001,
            AOFl = 7101,
            AOCd = 7201,
            Slew = 7601,
            AOSu = 7301,
            AOSe = 7401,
            AODi = 7501,
            GCLCt = 1351,
            GCLFl = 1391,
            MNm = 1211,
            Lg = 9999,
            //for expansion module
            ExW = 8001,
            ExBE = 8101,
            ExWE = 8201,
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
    public void StartTest()//iATester
    {
        textBox1.Text = dataHld.GetPara(Application.StartupPath);
        eStatus(this, new StatusEventArgs(iStatus.Running));
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
                var Obj01 = AdvantechHttpWebUtility.ParserJsonToObj<SysCoilData>(rawData);
                UpdateDevCoilStatus(Obj01);
                //
                ExeRes.Res = ExeCaseRes.Pass;Print(ExeRes);
                this.InvokeWaitStep();
                break;
            case ServiceAction.PatchSysInfo:
                break;
            case ServiceAction.GetNetConfig_ag:
                var Obj03 = AdvantechHttpWebUtility.ParserJsonToObj<SysCoilData>(rawData);
                UpdateDevCoilStatus(Obj03);
                //
                ExeRes.Res = ExeCaseRes.Pass; Print(ExeRes);
                this.InvokeWaitStep();
                break;
            //
            case ServiceAction.GetNetConfigforReg:
                var Obj04 = AdvantechHttpWebUtility.ParserJsonToObj<SysRegData>(rawData);
                UpdateDevRegStatus(Obj04);
                //
                ExeRes.Res = ExeCaseRes.Pass; Print(ExeRes);
                this.InvokeWaitStep();
                break;
            case ServiceAction.PatchSysInfoforReg:
                break;
            case ServiceAction.GetNetConfig_ag_forReg:
                var Obj05 = AdvantechHttpWebUtility.ParserJsonToObj<SysRegData>(rawData);
                UpdateDevRegStatus(Obj05);
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
        if (servAct == ServiceAction.GetNetConfig_ag_forReg)
        {
            VerifyRegItems();
        }
        else if (servAct == ServiceAction.PatchSysInfoforReg)
        {
            GetNetConfigRequestforReg();
        }
        else if (servAct == ServiceAction.GetNetConfigforReg)
        {
            PatchSysInfoRequestforReg();
        }
        else if (servAct == ServiceAction.GetNetConfig_ag)
        {
            VerifyCoilItems();
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
                                        "http://" + Device.IPAddress + "/modbus_coilconfig");
        //
        if (changeFlg) servAct = ServiceAction.GetNetConfig_ag;
        //
        ExeRes = new wResult()
        {
            Method = HttpRequestOption.GET,
            Ins = WISE_RESTFUL_URI.modbus_coilconfig,
        };
        
    }
    private void GetNetConfigRequestforReg()
    {
        Print(new wResult() { Des = "GetNetConfigRequest" });
        dataHld.SavePara(Application.StartupPath, textBox1.Text);
        Device.IPAddress = textBox1.Text;
        servAct = ServiceAction.GetNetConfigforReg;
        m_HttpRequest.SendGETRequest(Device.Account, Device.Password,
                                        "http://" + Device.IPAddress + "/modbus_regconfig");
        //
        if (changeFlg) servAct = ServiceAction.GetNetConfig_ag_forReg;
        //
        ExeRes = new wResult()
        {
            Method = HttpRequestOption.GET,
            Ins = WISE_RESTFUL_URI.modbus_regconfig,
        };

    }

    private void PatchSysInfoRequest()//Patch info
    {
        Print(new wResult() { Des = "PatchSysInfoRequest" });
        servAct = ServiceAction.PatchSysInfo;

        JavaScriptSerializer serializer = new JavaScriptSerializer();
        string sz_Jsonify = serializer.Serialize(ChangeDataArry);

        m_HttpRequest.SendPATCHRequest(Device.Account, Device.Password, GetURL(Device.IPAddress, Device.Port
                                    , WISE_RESTFUL_URI.modbus_coilconfig.ToString()), sz_Jsonify);
        changeFlg = true;
        //
        ExeRes = new wResult()
        {
            Method = HttpRequestOption.PATCH,
            Ins = WISE_RESTFUL_URI.modbus_coilconfig,
            Res = ExeCaseRes.Pass,
        }; Print(ExeRes);

        this.InvokeWaitStep();
    }
    private void PatchSysInfoRequestforReg()//Patch info
    {
        Print(new wResult() { Des = "PatchSysInfoRequest" });
        servAct = ServiceAction.PatchSysInfoforReg;

        JavaScriptSerializer serializer = new JavaScriptSerializer();
        string sz_Jsonify = serializer.Serialize(ChangeRegDataArry);

        m_HttpRequest.SendPATCHRequest(Device.Account, Device.Password, GetURL(Device.IPAddress, Device.Port
                                    , WISE_RESTFUL_URI.modbus_regconfig.ToString()), sz_Jsonify);
        changeFlg = true;
        //
        ExeRes = new wResult()
        {
            Method = HttpRequestOption.PATCH,
            Ins = WISE_RESTFUL_URI.modbus_regconfig,
            Res = ExeCaseRes.Pass,
        }; Print(ExeRes);

        this.InvokeWaitStep();
    }
    int errorCnt = 0;
    private void VerifyCoilItems()
    {
        Print(new wResult() { Des = "VerifyItems" });
        bool chk = false; changeFlg = false;

        if (GetDataArry.Md != ChangeDataArry.Md) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "Md  check is [" + GetDataArry.Md + "]", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        chk = false;
        if (GetDataArry.DI != ChangeDataArry.DI) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "DI  check is [" + GetDataArry.DI + "]", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        chk = false;
        if (GetDataArry.CtS != ChangeDataArry.CtS) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "CtS  check is [" + GetDataArry.CtS + "]", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        chk = false;
        if (GetDataArry.CtClr != ChangeDataArry.CtClr) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "CtClr  check is [" + GetDataArry.CtClr + "]", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        chk = false;
        if (GetDataArry.CtOv != ChangeDataArry.CtOv) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "CtOv  check is [" + GetDataArry.CtOv + "]", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        chk = false;
        if (GetDataArry.Lch != ChangeDataArry.Lch) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "Lch  check is [" + GetDataArry.Lch + "]", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        chk = false;
        if (GetDataArry.DO != ChangeDataArry.DO) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "DO  check is [" + GetDataArry.DO + "]", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        int tempErr = errorCnt;//<------ AI module cannot count.
        chk = false;
        if (GetDataArry.AIHR != ChangeDataArry.AIHR) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "AIHR  check is [" + GetDataArry.AIHR + "]", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        chk = false;
        if (GetDataArry.AILR != ChangeDataArry.AILR) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "AILR  check is [" + GetDataArry.AILR + "]", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        chk = false;
        if (GetDataArry.AIB != ChangeDataArry.AIB) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "AIB  check is [" + GetDataArry.AIB + "]", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        chk = false;
        if (GetDataArry.HAlm != ChangeDataArry.HAlm) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "HAlm  check is [" + GetDataArry.HAlm + "]", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        chk = false;
        if (GetDataArry.LAlm != ChangeDataArry.LAlm) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "LAlm  check is [" + GetDataArry.LAlm + "]", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        chk = false;
        //if (GetDataArry.GCtClr != ChangeDataArry.GCtClr) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "GCtClr  check is [" + GetDataArry.GCtClr + "]", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        // AI module cannot count.
        if (GetDataArry.AIHR != ChangeDataArry.AIHR && GetDataArry.AILR != ChangeDataArry.AILR)
            errorCnt = tempErr;
        //

        chk = false;
        if (GetDataArry.LB != ChangeDataArry.LB) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "LB  check is [" + GetDataArry.LB + "]", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        chk = false;
        //if (GetDataArry.ExB != ChangeDataArry.ExB) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "ExB  check is [" + GetDataArry.ExB + "]", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        //
        GetNetConfigRequestforReg();//<-------- Do next step
    }

    private void VerifyRegItems()
    {
        Print(new wResult() { Des = "VerifyItems" });
        bool chk = false; changeFlg = false;

        if (GetRegDataArry.Md != ChangeRegDataArry.Md) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "Md  check is [" + GetRegDataArry.Md + "]", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        chk = false;
        if (GetRegDataArry.DI != ChangeRegDataArry.DI) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "DI  check is [" + GetRegDataArry.DI + "]", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        chk = false;
        if (GetRegDataArry.CtFq != ChangeRegDataArry.CtFq) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "CtFq  check is [" + GetRegDataArry.CtFq + "]", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        chk = false;
        if (GetRegDataArry.DO != ChangeRegDataArry.DO) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "DO  check is [" + GetRegDataArry.DO + "]", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        chk = false;
        if (GetRegDataArry.PsLo != ChangeRegDataArry.PsLo) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "PsLo  check is [" + GetRegDataArry.PsLo + "]", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        chk = false;
        if (GetRegDataArry.PsHi != ChangeRegDataArry.PsHi) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "PsHi  check is [" + GetRegDataArry.PsHi + "]", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        chk = false;
        if (GetRegDataArry.PsAV != ChangeRegDataArry.PsAV) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "PsAV  check is [" + GetRegDataArry.PsAV + "]", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        chk = false;
        if (GetRegDataArry.PsIV != ChangeRegDataArry.PsIV) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "PsIV  check is [" + GetRegDataArry.PsIV + "]", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        int tempErr = errorCnt;//<------ AI module cannot count.
        chk = false;
        if (GetRegDataArry.AI != ChangeRegDataArry.AI) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "AI  check is [" + GetRegDataArry.AI + "]", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        chk = false;
        if (GetRegDataArry.HisH != ChangeRegDataArry.HisH) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "HisH  check is [" + GetRegDataArry.HisH + "]", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        chk = false;
        if (GetRegDataArry.HisL != ChangeRegDataArry.HisL) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "HisL  check is [" + GetRegDataArry.HisL + "]", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        chk = false;
        if (GetRegDataArry.AIF != ChangeRegDataArry.AIF) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "AIF  check is [" + GetRegDataArry.AIF + "]", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        chk = false;
        if (GetRegDataArry.HisHF != ChangeRegDataArry.HisHF) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "HisHF  check is [" + GetRegDataArry.HisHF + "]", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        chk = false;
        if (GetRegDataArry.HisLF != ChangeRegDataArry.HisLF) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "HisLF  check is [" + GetRegDataArry.HisLF + "]", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        chk = false;
        if (GetRegDataArry.AIFl != ChangeRegDataArry.AIFl) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "AIFl  check is [" + GetRegDataArry.AIFl + "]", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        chk = false;
        if (GetRegDataArry.AICd != ChangeRegDataArry.AICd) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "AICd  check is [" + GetRegDataArry.AICd + "]", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        chk = false;
        if (GetRegDataArry.AICh != ChangeRegDataArry.AICh) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "AICh  check is [" + GetRegDataArry.AICh + "]", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        chk = false;
        if (GetRegDataArry.AISc != ChangeRegDataArry.AISc) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "AISc  check is [" + GetRegDataArry.AISc + "]", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        chk = false;
        if (GetRegDataArry.AIPF != ChangeRegDataArry.AIPF) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "AIPF  check is [" + GetRegDataArry.AIPF + "]", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        // AO function is not work
        chk = false;
        //if (GetRegDataArry.AO != ChangeRegDataArry.AO) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "AO  check is [" + GetRegDataArry.AO + "]", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        chk = false;
        //if (GetRegDataArry.AOFl != ChangeRegDataArry.AOFl) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "AOFl  check is [" + GetRegDataArry.AOFl + "]", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        chk = false;
        //if (GetRegDataArry.AOCd != ChangeRegDataArry.AOCd) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "AOCd  check is [" + GetRegDataArry.AOCd + "]", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        chk = false;
        //if (GetRegDataArry.Slew != ChangeRegDataArry.Slew) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "Slew  check is [" + GetRegDataArry.Slew + "]", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        chk = false;
        //if (GetRegDataArry.AOSu != ChangeRegDataArry.AOSu) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "AOSu  check is [" + GetRegDataArry.AOSu + "]", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        chk = false;
        //if (GetRegDataArry.AOSe != ChangeRegDataArry.AOSe) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "AOSe  check is [" + GetRegDataArry.AOSe + "]", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        chk = false;
        //if (GetRegDataArry.AODi != ChangeRegDataArry.AODi) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "AODi  check is [" + GetRegDataArry.AODi + "]", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        chk = false;
        //if (GetRegDataArry.GCLCt != ChangeRegDataArry.GCLCt) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "GCLCt  check is [" + GetRegDataArry.GCLCt + "]", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        chk = false;
        //if (GetRegDataArry.GCLFl != ChangeRegDataArry.GCLFl) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "GCLFl  check is [" + GetRegDataArry.GCLFl + "]", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        //
        if (GetRegDataArry.AI != ChangeRegDataArry.AI && GetRegDataArry.HisH != ChangeRegDataArry.HisH)
            errorCnt = tempErr;


        //
        chk = false;
        if (GetRegDataArry.MNm != ChangeRegDataArry.MNm) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "MNm  check is [" + GetRegDataArry.MNm + "]", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        chk = false;
        if (GetRegDataArry.Lg != ChangeRegDataArry.Lg) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "Lg  check is [" + GetRegDataArry.Lg + "]", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        chk = false;
        //if (GetRegDataArry.ExW != ChangeRegDataArry.ExW) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "ExW  check is [" + GetRegDataArry.ExW + "]", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        chk = false;
        //if (GetRegDataArry.ExBE != ChangeRegDataArry.ExBE) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "ExBE  check is [" + GetRegDataArry.ExBE + "]", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass }); chk = false;
        chk = false;
        //if (GetRegDataArry.ExWE != ChangeRegDataArry.ExWE) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "ExWE  check is [" + GetRegDataArry.ExWE + "]", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });

        //Return the test result
        if (errorCnt > 0)
            eResult(this, new ResultEventArgs(iResult.Fail));
        else eResult(this, new ResultEventArgs(iResult.Pass));

        eStatus(this, new StatusEventArgs(iStatus.Completion));
    }

    #region ---- Update UI ----
    private void UpdateDevCoilStatus(SysCoilData data)
    {
        try
        {
            GetDataArry.Md = data.Md;
            GetDataArry.DI = data.DI;
            GetDataArry.CtS = data.CtS;
            GetDataArry.CtClr = data.CtClr;
            GetDataArry.CtOv = data.CtOv;
            GetDataArry.Lch = data.Lch;
            GetDataArry.DO = data.DO;
            GetDataArry.AIHR = data.AIHR;
            GetDataArry.AILR = data.AILR;
            GetDataArry.AIB = data.AIB;
            GetDataArry.HAlm = data.HAlm;
            GetDataArry.LAlm = data.LAlm;
            GetDataArry.GCtClr = data.GCtClr;
            GetDataArry.LB = data.LB;
            GetDataArry.ExB = data.ExB;

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
    private void UpdateDevRegStatus(SysRegData data)
    {
        try
        {
            GetRegDataArry.Md = data.Md;
            GetRegDataArry.DI = data.DI;
            GetRegDataArry.CtFq = data.CtFq;
            GetRegDataArry.PsLo = data.PsLo;
            GetRegDataArry.PsHi = data.PsHi;
            GetRegDataArry.PsAV = data.PsAV;
            GetRegDataArry.DO = data.DO;
            GetRegDataArry.PsIV = data.PsIV;
            GetRegDataArry.AI = data.AI;
            GetRegDataArry.HisH = data.HisH;
            GetRegDataArry.HisL = data.HisL;
            GetRegDataArry.AIF = data.AIF;
            GetRegDataArry.HisHF = data.HisHF;
            GetRegDataArry.HisLF = data.HisLF;
            GetRegDataArry.AIFl = data.AIFl;
            GetRegDataArry.AICd = data.AICd;
            GetRegDataArry.AICh = data.AICh;
            GetRegDataArry.AISc = data.AISc;
            GetRegDataArry.AIPF = data.AIPF;
            GetRegDataArry.AO = data.AO;
            GetRegDataArry.AOFl = data.AOFl;
            GetRegDataArry.AOCd = data.AOCd;
            GetRegDataArry.Slew = data.Slew;
            GetRegDataArry.AOSu = data.AOSu;
            GetRegDataArry.AOSe = data.AOSe;
            GetRegDataArry.AODi = data.AODi;
            GetRegDataArry.GCLCt = data.GCLCt;
            GetRegDataArry.GCLFl = data.GCLFl;
            GetRegDataArry.MNm = data.MNm;
            GetRegDataArry.Lg = data.Lg;
            GetRegDataArry.ExW = data.ExW;
            GetRegDataArry.ExBE = data.ExBE;
            GetRegDataArry.ExWE = data.ExWE;
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

    public class SysCoilData
    {
        public int Md { get; set; }
        public int DI { get; set; }
        public int CtS { get; set; }
        public int CtClr { get; set; }
        public int CtOv { get; set; }
        public int Lch { get; set; }
        public int DO { get; set; }
        public int AIHR { get; set; }
        public int AILR { get; set; }
        public int AIB { get; set; }
        public int HAlm { get; set; }
        public int LAlm { get; set; }
        public int GCtClr { get; set; }
        public int LB { get; set; }
        public int ExB { get; set; }
    }

    public class SysRegData
    {
        public int Md { get; set; }
        public int DI { get; set; }
        public int CtFq { get; set; }
        public int DO { get; set; }
        public int PsLo { get; set; }
        public int PsHi { get; set; }
        public int PsAV { get; set; }
        public int PsIV { get; set; }
        public int AI { get; set; }
        public int HisH { get; set; }
        public int HisL { get; set; }
        public int AIF { get; set; }
        public int HisHF { get; set; }
        public int HisLF { get; set; }
        public int AIFl { get; set; }
        public int AICd { get; set; }
        public int AICh { get; set; }
        public int AISc { get; set; }
        public int AIPF { get; set; }
        public int AO { get; set; }
        public int AOFl { get; set; }
        public int AOCd { get; set; }
        public int Slew { get; set; }
        public int AOSu { get; set; }
        public int AOSe { get; set; }
        public int AODi { get; set; }
        public int GCLCt { get; set; }
        public int GCLFl { get; set; }
        public int MNm { get; set; }
        public int Lg { get; set; }
        public int ExW { get; set; }
        public int ExBE { get; set; }
        public int ExWE { get; set; }
    }    

    public enum ServiceAction
    {
        Idel = 0,
        GetNetConfig = 1,
        PatchSysInfo = 2,
        GetNetConfig_ag = 3,
        GetNetConfigforReg = 4,
        PatchSysInfoforReg = 5,
        GetNetConfig_ag_forReg = 6,
        Verify = 7,

        Done = 99,
    }
}

