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
    bool GetRTUConfg = false;
    ServiceAction servAct = new ServiceAction();

    private delegate void DataGridViewCtrlAddDataRow(DataGridViewRow i_Row);
    private DataGridViewCtrlAddDataRow m_DataGridViewCtrlAddDataRow;
    internal const int Max_Rows_Val = 65535;

    SysData GetDataArry = new SysData();
    SysData ChangeDataArry = new SysData();//change description content
    //
    SysDataRTU GetDataArryRTU = new SysDataRTU();
    SysDataRTU ChangeDataArryRTU = new SysDataRTU();//change description content

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

        ChangeDataArry = new SysData()
        {
            BR = 7,
            DB = 1,
            P = 2,
            SB = 1,
            Prot = 0,
        };

        ChangeDataArryRTU = new SysDataRTU()
        {
            RT = 5000,
            DBP = 1000,
            EnC = 1,
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
                if (GetRTUConfg)
                {
                    var Obj02 = AdvantechHttpWebUtility.ParserJsonToObj<SysDataRTU>(rawData);
                    UpdateDevUIStatusRTU(Obj02);
                }
                else
                {
                    var Obj01 = AdvantechHttpWebUtility.ParserJsonToObj<SysData>(rawData);
                    UpdateDevUIStatus(Obj01);
                }
                
                //
                ExeRes.Res = ExeCaseRes.Pass;Print(ExeRes);
                this.InvokeWaitStep();
                break;
            case ServiceAction.PatchSysInfo:
                break;
            case ServiceAction.GetNetConfig_ag:
                if (GetRTUConfg)
                {
                    var Obj04 = AdvantechHttpWebUtility.ParserJsonToObj<SysDataRTU>(rawData);
                    UpdateDevUIStatusRTU(Obj04);
                }
                else
                {
                    var Obj03 = AdvantechHttpWebUtility.ParserJsonToObj<SysData>(rawData);
                    UpdateDevUIStatus(Obj03);
                }                
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
        if(GetRTUConfg)
        {
            Print(new wResult() { Des = "GetNetConfigRequest" });
            dataHld.SavePara(Application.StartupPath, textBox1.Text);
            Device.IPAddress = textBox1.Text;
            servAct = ServiceAction.GetNetConfig;
            m_HttpRequest.SendGETRequest(Device.Account, Device.Password,
                                            "http://" + Device.IPAddress + "/modbusslave_genconfig/com_1");
            //
            if (changeFlg) servAct = ServiceAction.GetNetConfig_ag;
        }
        else
        {
            Print(new wResult() { Des = "GetNetConfigRequest" });
            dataHld.SavePara(Application.StartupPath, textBox1.Text);
            Device.IPAddress = textBox1.Text;
            servAct = ServiceAction.GetNetConfig;
            m_HttpRequest.SendGETRequest(Device.Account, Device.Password,
                                            "http://" + Device.IPAddress + "/serial_config/com_1");
            //
            if (changeFlg) servAct = ServiceAction.GetNetConfig_ag;
        }

        ExeRes = new wResult()
        {
            Method = HttpRequestOption.GET,
            Ins = WISE_RESTFUL_URI.serial_config,
        };
    }

    private void PatchSysInfoRequest()//Patch info
    {
        Print(new wResult() { Des = "PatchSysInfoRequest" });
        if (GetRTUConfg)
        {
            servAct = ServiceAction.PatchSysInfo;

            JavaScriptSerializer serializer = new JavaScriptSerializer();
            string sz_Jsonify = serializer.Serialize(ChangeDataArryRTU);

            m_HttpRequest.SendPATCHRequest(Device.Account, Device.Password, GetURL(
                                        Device.IPAddress, Device.Port
                                        , WISE_RESTFUL_URI.modbusslave_genconfig.ToString() + "/com_1")
                                        , sz_Jsonify);
            changeFlg = true;
        }
        else
        {
            servAct = ServiceAction.PatchSysInfo;

            JavaScriptSerializer serializer = new JavaScriptSerializer();
            string sz_Jsonify = serializer.Serialize(ChangeDataArry);

            m_HttpRequest.SendPATCHRequest(Device.Account, Device.Password, GetURL(
                                        Device.IPAddress, Device.Port
                                        , WISE_RESTFUL_URI.serial_config.ToString() + "/com_1")
                                        , sz_Jsonify);
            changeFlg = true;          
        }

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
        int errorCnt = 0; changeFlg = false; bool chk = false;
        Print(new wResult() { Des = "VerifyItems" });        

        if (!GetRTUConfg)
        {
            if (GetDataArry.BR != ChangeDataArry.BR) { chk = true; errorCnt++; }
            Print(new wResult() { Des = "BR check is [" + GetDataArry.BR + "]", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
            chk = false;
            if (GetDataArry.DB != ChangeDataArry.DB) { chk = true; errorCnt++; }
            Print(new wResult() { Des = "DB check is [" + GetDataArry.DB + "]", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
            chk = false;
            if (GetDataArry.P != ChangeDataArry.P) { chk = true; errorCnt++; }
            Print(new wResult() { Des = "P check is [" + GetDataArry.P + "]", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
            chk = false;
            if (GetDataArry.SB != ChangeDataArry.SB) { chk = true; errorCnt++; }
            Print(new wResult() { Des = "SB check is [" + GetDataArry.SB + "]", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
            chk = false;
            if (GetDataArry.Prot != ChangeDataArry.Prot) { chk = true; errorCnt++; }
            Print(new wResult() { Des = "Prot check is [" + GetDataArry.Prot + "]", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });

            GetRTUConfg = true;
            GetNetConfigRequest();
        }
        else
        {
            if (GetDataArry.BR != ChangeDataArry.BR) { chk = true; errorCnt++; }
            Print(new wResult() { Des = "BR check is [" + GetDataArry.BR + "]", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
            chk = false;
            if (GetDataArry.DB != ChangeDataArry.DB) { chk = true; errorCnt++; }
            Print(new wResult() { Des = "DB check is [" + GetDataArry.DB + "]", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });
            chk = false;
            if (GetDataArry.P != ChangeDataArry.P) { chk = true; errorCnt++; }
            Print(new wResult() { Des = "P check is [" + GetDataArry.P + "]", Res = chk ? ExeCaseRes.Fail : ExeCaseRes.Pass });

            GetRTUConfg = false;

            //Return the test result
            if (errorCnt > 0)
                eResult(this, new ResultEventArgs(iResult.Fail));
            else eResult(this, new ResultEventArgs(iResult.Pass));

            eStatus(this, new StatusEventArgs(iStatus.Completion));
        }
        
    }

    #region ---- Update UI ----
    private void UpdateDevUIStatus(SysData data)
    {
        try
        {
            GetDataArry.BR = data.BR;
            GetDataArry.DB = data.DB;
            GetDataArry.P = data.P;
            GetDataArry.SB = data.SB;
            GetDataArry.Prot = data.Prot;
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
    private void UpdateDevUIStatusRTU(SysDataRTU data)
    {
        try
        {
            GetDataArryRTU.RT = data.RT;
            GetDataArryRTU.DBP = data.DBP;
            GetDataArryRTU.EnC = data.EnC;
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
        public int BR { get; set; }
        public int DB { get; set; }
        public int P { get; set; }
        public int SB { get; set; }
        public int Prot { get; set; }
    }
    public class SysDataRTU
    {
        public int RT { get; set; }
        public int DBP { get; set; }
        public int EnC { get; set; }
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

