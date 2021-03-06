﻿using System;
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
    bool DevConnFail = false;
    ServiceAction servAct = new ServiceAction();

    private delegate void DataGridViewCtrlAddDataRow(DataGridViewRow i_Row);
    private DataGridViewCtrlAddDataRow m_DataGridViewCtrlAddDataRow;
    internal const int Max_Rows_Val = 65535;

    SysData GetDataArry = new SysData();
    SysChgData ChangeDataArry = new SysChgData();//change description content
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

        ChangeDataArry = new SysChgData()
        {
            Nm = "123456789012345678901234567890123456789012345678901234567890ABC",
            Des = "123456789012345678901234567890123456789012345678901234567890"
                    + "123456789012345678901234567890123456789012345678901234567890"
                    + "ABCDEFG",
            UID = "123456789012345678901234567890A",
            Loc = "123456789012345678901234567890123456789012345678901234567890"
                    + "123456789012345678901234567890123456789012345678901234567890"
                    + "ABCDEFG",
            Lat = "1234567890ABCDE",
            Lng = "1234567890ABCDE",
            Alt = "1234567890ABCDE",
            Tm = "2000-12-31T23:59:59-05:00",

        };
    }

    private void Form1_Load(object sender, EventArgs e)
    {
        //Get IO value
        m_HttpRequest = new AdvantechHttpWebUtility();
        m_HttpRequest.ResquestOccurredError += this.OnGetHttpRequestError;
        m_HttpRequest.ResquestResponded += this.OnGetData;

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
        //button1.Text = HttpReq_Connet() ? "Connected" : "Disconnected";
    }
    public void StartTest()//iATester
    {
        eStatus(this, new StatusEventArgs(iStatus.Running));
        button1.Text = HttpReq_Connet() ? "Connected" : "Disconnected";
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
        DevConnFail = true;
        //
        ExeRes.Res = ExeCaseRes.Fail;
        ExeRes.Err = ex.ToString();
        Print(ExeRes);
    }

    private void OnGetData(string rawData)//Feedback Http request
    {
        if (!Device.ConnFlg) servAct = ServiceAction.Connection;//Priority High

        switch (servAct)
        {
            case ServiceAction.Connection:
                var Obj01 = AdvantechHttpWebUtility.ParserJsonToObj<GetDeviceData>(rawData);
                UpdateDevUIStatus(Obj01);
                //
                ExeRes.Res = ExeCaseRes.Pass;Print(ExeRes);
                this.InvokeWaitStep();
                break;
            case ServiceAction.GetSysInfo:
                var Obj02 = AdvantechHttpWebUtility.ParserJsonToObj<SysData>(rawData);
                UpdateSysInfoStatus(Obj02);
                //
                ExeRes.Res = ExeCaseRes.Pass; Print(ExeRes);
                this.InvokeWaitStep();
                break;
            case ServiceAction.PatchSysInfo:
                break;
            case ServiceAction.GetSysInfo_ag:
                var Obj03 = AdvantechHttpWebUtility.ParserJsonToObj<SysData>(rawData);
                UpdateSysInfoStatus(Obj03);
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
        button1.Text = HttpReq_Connet() ? "Connected" : "Disconnected";
    }

    private bool HttpReq_Connet()
    {
        bool res = false; changeFlg = false;
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

        GetDevInfoRequest();//實作HttpReq connect
        //RefreshMBList = false;
        //response to Device.IPAddress event
        int time_out = 0;
        while (!Device.ConnFlg)//propose is delay to wait response event.
        {
            if (DevConnFail)
            { DevConnFail = false; res = false; break; }
            else res = true;
            servAct = ServiceAction.Connection;
            System.Threading.Thread.Sleep(1000);
            time_out++;
            if (time_out > 4) { res = false; break; }
        }

        typTxt.Text = Device.ModuleType;
        fwTxt.Text = Device.FirmwareVer;
        if (!dataHld.SavePara(Application.StartupPath, textBox1.Text))
            MessageBox.Show("Save para fail.");
        Device.IPAddress = textBox1.Text;
        return res;
    }

    private void InvokeWaitStep()//start to read IO invoke
    {
        int m_iPollingTime = 500;
        Thread.Sleep(m_iPollingTime);//m_iPollingTime
        NextStep();
    }

    private void NextStep()
    {
        if (servAct == ServiceAction.GetSysInfo_ag)
        {
            VerifyItems();
        }
        else if (servAct == ServiceAction.PatchSysInfo)
        {
            GetSysInfoRequest();
        }
        else if (servAct == ServiceAction.GetSysInfo)
        {
            PatchSysInfoRequest();
        }
        else if (servAct == ServiceAction.Connection)
        {
            GetSysInfoRequest();
        }
    }

    //Request Cmd
    private void GetDevInfoRequest()
    {
        Print(new wResult() { Des = "GetDevInfoRequest" });
        servAct = ServiceAction.Connection;
        m_HttpRequest.SendGETRequest(Device.Account, Device.Password,
                                        "http://" + Device.IPAddress + "/profile");
        //
        ExeRes = new wResult()
        {
            Method = HttpRequestOption.GET,
            Ins = WISE_RESTFUL_URI.profile,
        };
        
    }

    private void GetSysInfoRequest()//Get info
    {
        Print(new wResult() { Des = "GetSysInfo" });
        servAct = ServiceAction.GetSysInfo;
        m_HttpRequest.SendGETRequest(Device.Account, Device.Password,
                                        "http://" + Device.IPAddress + "/sys_info");
        //
        if (changeFlg) servAct = ServiceAction.GetSysInfo_ag;
        //
        ExeRes = new wResult()
        {
            Method = HttpRequestOption.GET,
            Ins = WISE_RESTFUL_URI.sys_info,
        };
    }

    private void PatchSysInfoRequest()//Patch info
    {
        Print(new wResult() { Des = "PatchSysInfoRequest" });
        servAct = ServiceAction.PatchSysInfo;

        JavaScriptSerializer serializer = new JavaScriptSerializer();
        string sz_Jsonify = serializer.Serialize(ChangeDataArry);

        m_HttpRequest.SendPATCHRequest(Device.Account, Device.Password, GetURL(Device.IPAddress, Device.Port
                                    , WISE_RESTFUL_URI.sys_info.ToString()), sz_Jsonify);
        changeFlg = true;
        //
        ExeRes = new wResult()
        {
            Method = HttpRequestOption.PATCH,
            Ins = WISE_RESTFUL_URI.sys_info,
            Res = ExeCaseRes.Pass,
        }; Print(ExeRes);

        this.InvokeWaitStep();
    }

    private void VerifyItems()
    {
        int errorCnt = 0; changeFlg = false;
        Print(new wResult() { Des = "VerifyItems" });
        bool chk = false;
        if (GetDataArry.Nm != ChangeDataArry.Nm) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "Nm check", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        chk = false;
        if (GetDataArry.Des != ChangeDataArry.Des) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "Des check", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        chk = false;
        if (GetDataArry.UID != ChangeDataArry.UID) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "UID check", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        chk = false;
        if (GetDataArry.Loc != ChangeDataArry.Loc) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "Loc check", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        chk = false;
        if (GetDataArry.Lat != ChangeDataArry.Lat) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "Lat check", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        chk = false;
        if (GetDataArry.Lng != ChangeDataArry.Lng) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "Lng check", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        chk = false;
        if (GetDataArry.Alt != ChangeDataArry.Alt) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "Alt check", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        //chk = false;
        //if (GetDataArry.Tm != ChangeDataArry.Tm) { chk = true; errorCnt++; }
        //Print(new wResult() { Des = "Tm check", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });


        //Return the test result
        if (errorCnt > 0)
            eResult(this, new ResultEventArgs(iResult.Fail));
        else eResult(this, new ResultEventArgs(iResult.Pass));

        eStatus(this, new StatusEventArgs(iStatus.Completion));
    }

    #region ---- Update UI ----
    private void UpdateDevUIStatus(GetDeviceData devdata)
    {
        try
        {
            if (devdata.Dev != null)
            {
                for (int i = 0; i < devdata.Dev.Length; ++i)
                {   //將資料存入Device
                    Device.ModuleType = devdata.Dev[i].Id.ToString();
                    Device.FirmwareVer = devdata.Dev[i].FwVer.ToString();
                    //Device.Account = devdata.Dev[i]..ToString();
                    //Device.Password = devdata.Dev[i].FwVer.ToString();
                    Device.DiTotal = devdata.Dev[i].DIn;
                    Device.DoTotal = devdata.Dev[i].DOn;
                    Device.RLaTotal = devdata.Dev[i].RLAn;
                    Device.AiTotal = devdata.Dev[i].AIn;
                    Device.UInTotal = devdata.Dev[i].UIn;
                    Device.AoTotal = devdata.Dev[i].AOn;
                    Device.CntTotal = devdata.Dev[i].Cntn;
                    Device.ConnFlg = true;
                }
            }
            else
            {
                throw new Exception("Parser Conn Data Fail");
            }
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

    private void UpdateSysInfoStatus(SysData data)
    {
        try
        {
            GetDataArry.Id = data.Id;
            GetDataArry.Nm = data.Nm;
            GetDataArry.Des = data.Des;
            GetDataArry.UID = data.UID;
            GetDataArry.Loc = data.Loc;
            GetDataArry.Lat = data.Lat;
            GetDataArry.Lng = data.Lng;
            GetDataArry.Alt = data.Alt;
            GetDataArry.Tm = data.Tm;
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
        public string Id { get; set; }
        public string Nm { get; set; }
        public string Des { get; set; }
        public string UID { get; set; }
        public string Loc { get; set; }
        public string Lat { get; set; }
        public string Lng { get; set; }
        public string Alt { get; set; }
        public string Tm { get; set; }
    }

    public class SysChgData
    {
        public string Nm { get; set; }
        public string Des { get; set; }
        public string UID { get; set; }
        public string Loc { get; set; }
        public string Lat { get; set; }
        public string Lng { get; set; }
        public string Alt { get; set; }
        public string Tm { get; set; }
    }

    //public class wResult//Result class
    //{
    //    public HttpRequestOption Method//Method
    //    {
    //        get;
    //        set;
    //    }
    //    public WISE_RESTFUL_URI Ins//instruction
    //    {
    //        get;
    //        set;
    //    }
    //    public string Des//description
    //    {
    //        get;
    //        set;
    //    }
    //    public ExeCaseRes Res//Result
    //    {
    //        get;
    //        set;
    //    }
    //    public string Err//Error code
    //    {
    //        get;
    //        set;
    //    }
    //}

    public enum ServiceAction
    {
        Idel = 0,
        Connection = 1,
        GetSysInfo = 2,
        PatchSysInfo = 3, 
        GetSysInfo_ag = 4,
        Verify = 5,

        Done = 99,
    }
}

