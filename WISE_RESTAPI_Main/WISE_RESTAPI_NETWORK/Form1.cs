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
    SysChgData ChangeInfraDataArry = new SysChgData();//When in infra mode
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
                                        "http://" + Device.IPAddress + "/net_basic");
        //
        if (changeFlg) servAct = ServiceAction.GetNetConfig_ag;
        //
        ExeRes = new wResult()
        {
            Method = HttpRequestOption.GET,
            Ins = WISE_RESTFUL_URI.net_basic,
        };
        
    }

    private void PatchSysInfoRequest()//Patch info
    {
        Print(new wResult() { Des = "PatchSysInfoRequest" });
        servAct = ServiceAction.PatchSysInfo;

        JavaScriptSerializer serializer = new JavaScriptSerializer();
        string sz_Jsonify = serializer.Serialize(GetChangeDataArray());

        m_HttpRequest.SendPATCHRequest(Device.Account, Device.Password, GetURL(Device.IPAddress, Device.Port
                                    , WISE_RESTFUL_URI.net_basic.ToString()), sz_Jsonify);
        changeFlg = true;
        //
        ExeRes = new wResult()
        {
            Method = HttpRequestOption.PATCH,
            Ins = WISE_RESTFUL_URI.net_basic,
            Res = ExeCaseRes.Pass,
        }; Print(ExeRes);

        this.InvokeWaitStep();
    }

    private void VerifyItems()
    {
        int errorCnt = 0; changeFlg = false;
        Print(new wResult() { Des = "VerifyItems" });
        bool chk = false;
        Print(new wResult() { Des = "MAC check [" + GetDataArry.MAC + "]", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        if (GetDataArry.DHCP != GetChangeDataArray().DHCP) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "DHCP check", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        chk = false;
        if (GetDataArry.IP != GetChangeDataArray().IP || GetDataArry.IP == null) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "IP check [" + GetDataArry.IP + "]", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        chk = false;
        if (GetDataArry.Msk != GetChangeDataArray().Msk || GetDataArry.Msk == null) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "Msk check [" + GetDataArry.Msk + "]", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
        chk = false;
        if (GetDataArry.GW != GetChangeDataArray().GW || GetDataArry.GW == null) { chk = true; errorCnt++; }
        Print(new wResult() { Des = "GW check [" + GetDataArry.GW + "]", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });

        //Return the test result
        if (errorCnt > 0)
            eResult(this, new ResultEventArgs(iResult.Fail));
        else eResult(this, new ResultEventArgs(iResult.Pass));

        eStatus(this, new StatusEventArgs(iStatus.Completion));
    }
    private SysChgData GetChangeDataArray()
    {
        ChangeDataArry = new SysChgData()
        {
            DHCP = 1,
            IP = Device.IPAddress,
            Msk = "255.255.255.0",
            GW = "192.168.0.1",
        };
        ChangeInfraDataArry = new SysChgData()
        {
            DHCP = 0,
            IP = Device.IPAddress,
            Msk = "255.255.0.0",
            GW = "192.168.0.1",
        };

        if (chkMod.Checked)
            return ChangeDataArry;
        else return ChangeInfraDataArry;
    }
    //private string SetDevIP()
    //{
    //    string ip = "";
    //    if (Device.ModuleType == "WISE-4050")
    //        ip = "192.168.0.66";
    //    else if (Device.ModuleType == "WISE-4060")
    //        ip = "192.168.0.67";
    //    else if (Device.ModuleType == "WISE-4012")
    //        ip = "192.168.0.68";
    //    else if (Device.ModuleType == "WISE-4012E")
    //        ip = "192.168.0.69";
    //    else if (Device.ModuleType == "WISE-4051")
    //        ip = "192.168.0.70";
    //    else
    //        ip = "192.168.0.99";

    //    return ip;
    //}

    #region ---- Update UI ----
    private void UpdateDevUIStatus(SysData data)
    {
        try
        {
            GetDataArry.MAC = data.MAC;
            GetDataArry.DHCP = data.DHCP;
            GetDataArry.IP = data.IP;
            GetDataArry.Msk = data.Msk;
            GetDataArry.GW = data.GW;
            
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
        public string MAC { get; set; }
        public int DHCP { get; set; }
        public string IP { get; set; }
        public string Msk { get; set; }
        public string GW { get; set; }
    }

    public class SysChgData
    {
        public int DHCP { get; set; }
        public string IP { get; set; }
        public string Msk { get; set; }
        public string GW { get; set; }
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

