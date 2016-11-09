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

    Acc[] GetDataArry = new Acc[8];
    //SysChgData[] ChangeDataArry = new SysChgData[8];//change description content
    bool changeFlg = false;
    wResult ExeRes;
    int indx = 0;
    string[] IpTable = new string[8];

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

        GetDataArry.Initialize();
        //ChangeDataArry.Initialize();
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
        //
        // 取得本機名稱
        string strHostName = Dns.GetHostName();
        // 取得本機的IpHostEntry類別實體，MSDN建議新的用法
        IPHostEntry iphostentry = Dns.GetHostEntry(strHostName);
        // 取得所有 IP 位址
        System.Collections.ArrayList ipList = new System.Collections.ArrayList();
        foreach (IPAddress ipaddress in iphostentry.AddressList)
        {
            // 只取得IP V4的Address
            if (ipaddress.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork)
            {
                ipList.Add(ipaddress.ToString());
                //Console.WriteLine("Local IP: " + ipaddress.ToString());
            }
        }
        //input
        //foreach (var item in ipList)
        //{
        //    ChangeDataArry[indx] = new SysChgData() { En = 1, Adr = (string)item };
        //    indx++;
        //}
        //for (int i = indx; i < 8; i++)
        //{
        //    ChangeDataArry[i] = new SysChgData() { En = 1, Adr = "192.168.1." + (i + 1).ToString() };
        //}
        IpTable.Initialize();
        foreach (var item in ipList)
        {
            IpTable[indx] = (string)item;
            indx++;
        }
        for (int i = indx; i < 8; i++)
            IpTable[i] = "192.168.1." + (i + 1).ToString();
    }
    public void StartTest()//iATester
    {
        textBox1.Text = dataHld.GetPara(Application.StartupPath);
        indx = 0;
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
                var Obj01 = AdvantechHttpWebUtility.ParserJsonToObj<Acc>(rawData);
                UpdateDevUIStatus(Obj01);
                //
                ExeRes.Res = ExeCaseRes.Pass;Print(ExeRes);
                this.InvokeWaitStep();
                break;
            case ServiceAction.PatchSysInfo:
                break;
            case ServiceAction.GetNetConfig_ag:
                var Obj03 = AdvantechHttpWebUtility.ParserJsonToObj<Acc>(rawData);
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
        indx = 0;
        dataHld.SavePara(Application.StartupPath, textBox1.Text);
        Device.IPAddress = textBox1.Text;
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
        indx++;
        if (indx > 23)
        {
            VerifyItems();
        }
        else if (indx > 15)
        {
            GetNetConfigRequest();
        }
        else if (indx > 7)
        {
            PatchSysInfoRequest(); 
        }
        else if (indx < 8)
        {
            GetNetConfigRequest();
        }
        
        
    }

    //Request Cmd
    private void GetNetConfigRequest()
    {
        Print(new wResult() { Des = "GetNetConfigRequest" });
        //dataHld.SavePara(Application.StartupPath, textBox1.Text);
        //Device.IPAddress = textBox1.Text;
        servAct = ServiceAction.GetNetConfig;

        int temp_idx = indx > 14 ? indx - 16 : indx;

        m_HttpRequest.SendGETRequest(Device.Account, Device.Password, GetURL(Device.IPAddress, Device.Port
                                    , WISE_RESTFUL_URI.accessctrl.ToString()) + "/idx_" + temp_idx.ToString());
        //
        if (changeFlg) 
            servAct = ServiceAction.GetNetConfig_ag;
        //
        ExeRes = new wResult()
        {
            Method = HttpRequestOption.GET,
            Ins = WISE_RESTFUL_URI.accessctrl,
            Des = "/idx_" + temp_idx.ToString(),
        };
        
    }

    private void PatchSysInfoRequest()//Patch info
    {
        Print(new wResult() { Des = "PatchSysInfoRequest" });
        servAct = ServiceAction.PatchSysInfo;

        int temp_idx = indx - 8;

        JavaScriptSerializer serializer = new JavaScriptSerializer();
        //string sz_Jsonify = serializer.Serialize(ChangeDataArry[temp_idx]);
        string sz_Jsonify = serializer.Serialize(new SysChgData { En = 1, Adr = IpTable[temp_idx] });

        m_HttpRequest.SendPATCHRequest(Device.Account, Device.Password, GetURL(Device.IPAddress, Device.Port
                                    , WISE_RESTFUL_URI.accessctrl.ToString() + "/idx_" + temp_idx.ToString()), sz_Jsonify);
        //
        ExeRes = new wResult()
        {
            Method = HttpRequestOption.PATCH,
            Ins = WISE_RESTFUL_URI.accessctrl,
            Res = ExeCaseRes.Pass,
            Des = "/idx_" + temp_idx.ToString(),
        }; Print(ExeRes);

        this.InvokeWaitStep();
    }

    private void VerifyItems()
    {
        int errorCnt = 0; changeFlg = false;
        Print(new wResult() { Des = "VerifyItems" });
        for (int i = 0; i < 8; i++)
        {
            bool chk = false;
            if (GetDataArry[i].En != 1) { chk = true; errorCnt++; }
            Print(new wResult()
            {
                Des = "En check by idx[" + i.ToString() + "]"
                ,
                Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass
            });
            chk = false;
            if (GetDataArry[i].Adr != IpTable[i] || GetDataArry[i].Adr == null) { chk = true; errorCnt++; }
            Print(new wResult()
            {
                Des = "Adr check by idx[" + i.ToString() + "]"
                ,
                Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass
            });
        }

        //chk = false;



        //Return the test result
        if (errorCnt > 0)
            eResult(this, new ResultEventArgs(iResult.Fail));
        else eResult(this, new ResultEventArgs(iResult.Pass));

        eStatus(this, new StatusEventArgs(iStatus.Completion));
        //if (errorCnt > 0) ias.ReturnRes((int)iALibrary.iResult.Fail);
        //else ias.ReturnRes((int)iALibrary.iResult.Pass);
        ////To notify iATester the test is completion
        //ias.iStatus(iALibrary.iStatus.Completion);
    }

    #region ---- Update UI ----
    private void UpdateDevUIStatus(Acc data)
    {
        try
        {
            int temp_idx = indx > 15 ? indx - 16 : indx;
            GetDataArry[temp_idx] = new Acc();
            GetDataArry[temp_idx].En = data.En;
            GetDataArry[temp_idx].Adr = data.Adr;            
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
    public class Acc
    {
        public int Idx { get; set; }
        public int En { get; set; }
        public string Adr { get; set; }
    }

    public class SysChgData
    {
        public int En { get; set; }
        public string Adr { get; set; }
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

