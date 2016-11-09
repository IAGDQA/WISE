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
        if(!DIModChkBlock())
        {
            PrintTitle("Module not support.");
            return;
        }
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
        if (dev.ModuleType.ToUpper() == "WISE-4012")
        {
            UIConfgBlock();
            UDIExeDIconfig();
            UDIExeCntconfig();
            UDIExeL2Hconfig();
            UDIExeH2Lconfig();
        }
        else
        {
            DIExeDIconfig();
            DIExeCntconfig();
            DIExeL2Hconfig();
            DIExeH2Lconfig();
            DIExeFreqconfig();
        }
            
             

        selenium.Close();
        
    }

    int di_num = 0;
    bool DIModChkBlock()
    {
        if (dev.ModuleType.ToUpper() == "WISE-4012E")
        {
            di_num = 2; return true;
        }
        else if (dev.ModuleType.ToUpper() == "WISE-4051")
        {
            di_num = 8; return true;
        }
        else if (dev.ModuleType.ToUpper() == "WISE-4050"
                    || dev.ModuleType.ToUpper() == "WISE-4060"
                    || dev.ModuleType.ToUpper() == "WISE-4012"
                    || dev.ModuleType.ToUpper() == "WISE-4050/LAN"
                    || dev.ModuleType.ToUpper() == "WISE-4060/LAN")
        {
            di_num = 4; return true;
        }
        return false;
    }
    void UIConfgBlock()//for WISE-4012
    {
        PrintTitle("UIConfg");
        selenium.Click("id=ioStatus0");
        System.Threading.Thread.Sleep(1000);
        for (int i = 0; i < di_num; i++)
        {
            selenium.Select("id=selMd_" + i.ToString(), "label=DI");
        }
        selenium.Click("id=btnUIConfig");
        PrintStep();
    }
    void DIExeDIconfig()
    {
        PrintTitle("DIConfgSts");
        selenium.Click("id=ioStatus0");
        if (dev.ModuleType.ToUpper() == "WISE-4012E") selenium.Click("link=DI");
        selenium.Click("id=diConfigRow0");
        if (dev.ModuleType.ToUpper() == "WISE-4012E")
            selenium.Click("xpath=(//a[contains(text(),'Configuration')])[2]");
        else selenium.Click("//a[contains(text(),'Configuration')]");
        PrintStep();
        System.Threading.Thread.Sleep(1000);
        //
        for (int i = 0; i < di_num; i++)
        {
            if (dev.ModuleType.ToUpper() == "WISE-4012E")
            {
                selenium.Select("css=#diConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-3 > div.input-group > #selCh", "label=" + i.ToString());
                selenium.Type("css=#diConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-6 > div.input-group > #inpTag", "ABCDEFGHIJKLMNOPQRST" + i.ToString());
            }
            else
            {
                selenium.Select("id=selCh", "label=" + i.ToString());
                selenium.Type("id=inpTag", "ABCDEFGHIJKLMNOPQRST" + i.ToString());
            }
                
            selenium.Select("id=selMd", "label=DI");
            selenium.Click("id=inpInv");
            selenium.Click("id=inpFltr");
            selenium.Type("id=inpFtLo", "65535");
            selenium.Type("id=inpFtHi", "65535");
            if (dev.ModuleType.ToUpper() == "WISE-4012E")
                selenium.Click("css=#diConfigBaseForm > div.panel-footer.clearfix > div.pull-right > #btnSubmit");
            else
                selenium.Click("id=btnSubmit");
            System.Threading.Thread.Sleep(1000); PrintStep();
        }
        //for all
        if (dev.ModuleType.ToUpper() == "WISE-4012E")
            selenium.Select("css=#diConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-3 > div.input-group > #selCh", "label=All");
        else
            selenium.Select("id=selCh", "label=All");
        selenium.Select("id=selMd", "label=DI");
        selenium.Click("id=inpInv");
        selenium.Click("id=inpFltr");
        selenium.Type("id=inpFtLo", "65535");
        selenium.Type("id=inpFtHi", "65535");
        if (dev.ModuleType.ToUpper() == "WISE-4012E")
            selenium.Click("css=#diConfigBaseForm > div.panel-footer.clearfix > div.pull-right > #btnSubmit");
        else
            selenium.Click("id=btnSubmit");
        System.Threading.Thread.Sleep(2000);
        //
        PrintStep();

    }

    void DIExeCntconfig()
    {
        PrintTitle("DIExeCntconfig");
        selenium.Click("id=ioStatus0");
        if (dev.ModuleType.ToUpper() == "WISE-4012E") selenium.Click("link=DI");
        selenium.Click("id=diConfigRow0");
        if (dev.ModuleType.ToUpper() == "WISE-4012E")
            selenium.Click("xpath=(//a[contains(text(),'Configuration')])[2]");
        else selenium.Click("//a[contains(text(),'Configuration')]");
        PrintStep();
        //
        for (int i = 0; i < di_num; i++)
        {
            if (dev.ModuleType.ToUpper() == "WISE-4012E")
            {
                selenium.Select("css=#diConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-3 > div.input-group > #selCh", "label=" + i.ToString());
                selenium.Type("css=#diConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-6 > div.input-group > #inpTag", "ABCDEFGHIJKLMNOPQRS" + (i + 10).ToString());
            }
            else
            {
                selenium.Select("id=selCh", "label=" + i.ToString());
                selenium.Type("id=inpTag", "ABCDEFGHIJKLMNOPQRS" + (i + 10).ToString());
            }
            
            selenium.Select("id=selMd", "label=Counter");
            selenium.Click("id=inpInv");
            selenium.Click("id=inpFltr");
            selenium.Type("id=inpFtLo", "65535");
            selenium.Type("id=inpFtHi", "65535");
            selenium.Type("id=inpCntIV", "4294967295");
            selenium.Click("id=inpCntKp");

            if (dev.ModuleType.ToUpper() == "WISE-4012E")
                selenium.Click("css=#diConfigBaseForm > div.panel-footer.clearfix > div.pull-right > #btnSubmit");
            else
                selenium.Click("id=btnSubmit");
            System.Threading.Thread.Sleep(1000); PrintStep();
        }
        //for all
        if (dev.ModuleType.ToUpper() == "WISE-4012E")
            selenium.Select("css=#diConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-3 > div.input-group > #selCh", "label=All");
        else
            selenium.Select("id=selCh", "label=All");
        selenium.Select("id=selMd", "label=Counter");
        selenium.Click("id=inpInv");
        selenium.Click("id=inpFltr");
        selenium.Type("id=inpFtLo", "65535");
        selenium.Type("id=inpFtHi", "65535");
        selenium.Type("id=inpCntIV", "4294967295");
        selenium.Click("id=inpCntKp");

        if (dev.ModuleType.ToUpper() == "WISE-4012E")
            selenium.Click("css=#diConfigBaseForm > div.panel-footer.clearfix > div.pull-right > #btnSubmit");
        else
            selenium.Click("id=btnSubmit");
        System.Threading.Thread.Sleep(1000);
        //
        PrintStep();
        //
        PrintTitle("Check DI Counter View");
        selenium.Click("link=Status");
        
        for (int i = 0; i < di_num; i++)
        {
            selenium.Click("id=switchonOffSwitchDI_" + i.ToString());
            selenium.Click("id=btnResetDI_" + i.ToString());
            selenium.Click("id=switchonOffSwitchDI_" + i.ToString());
            PrintStep();
        }
        
    }

    void DIExeL2Hconfig()
    {
        PrintTitle("DIExeL2Hconfig");
        selenium.Click("id=ioStatus0");
        if (dev.ModuleType.ToUpper() == "WISE-4012E") selenium.Click("link=DI");
        selenium.Click("id=diConfigRow0");
        if (dev.ModuleType.ToUpper() == "WISE-4012E")
            selenium.Click("xpath=(//a[contains(text(),'Configuration')])[2]");
        else selenium.Click("//a[contains(text(),'Configuration')]");
        PrintStep();
        //
        for (int i = 0; i < di_num; i++)
        {
            if (dev.ModuleType.ToUpper() == "WISE-4012E")
            {
                selenium.Select("css=#diConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-3 > div.input-group > #selCh", "label=" + i.ToString());
                selenium.Type("css=#diConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-6 > div.input-group > #inpTag", "ABCDEFGHIJKLMNOPQRS" + (i + 20).ToString());
            }
            else
            {
                selenium.Select("id=selCh", "label=" + i.ToString());
                selenium.Type("id=inpTag", "ABCDEFGHIJKLMNOPQRS" + (i + 20).ToString());
            }

            selenium.Select("id=selMd", "label=Low to High Latch");
            selenium.Click("id=inpInv");
            if (dev.ModuleType.ToUpper() == "WISE-4012E")
                selenium.Click("css=#diConfigBaseForm > div.panel-footer.clearfix > div.pull-right > #btnSubmit");
            else
                selenium.Click("id=btnSubmit");
            System.Threading.Thread.Sleep(1000); PrintStep();
        }
        //for all
        if (dev.ModuleType.ToUpper() == "WISE-4012E")
            selenium.Select("css=#diConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-3 > div.input-group > #selCh", "label=All");
        else
            selenium.Select("id=selCh", "label=All");
        selenium.Select("id=selMd", "label=Low to High Latch");
        selenium.Click("id=inpInv");
        if (dev.ModuleType.ToUpper() == "WISE-4012E")
            selenium.Click("css=#diConfigBaseForm > div.panel-footer.clearfix > div.pull-right > #btnSubmit");
        else
            selenium.Click("id=btnSubmit");
        System.Threading.Thread.Sleep(1000);
        //
        PrintStep();

    }

    void DIExeH2Lconfig()
    {
        PrintTitle("DIExeH2Lconfig");
        selenium.Click("id=ioStatus0");
        if (dev.ModuleType.ToUpper() == "WISE-4012E") selenium.Click("link=DI");
        selenium.Click("id=diConfigRow0");
        if (dev.ModuleType.ToUpper() == "WISE-4012E")
            selenium.Click("xpath=(//a[contains(text(),'Configuration')])[2]");
        else selenium.Click("//a[contains(text(),'Configuration')]");
        PrintStep();
        //
        for (int i = 0; i < di_num; i++)
        {
            if (dev.ModuleType.ToUpper() == "WISE-4012E")
            {
                selenium.Select("css=#diConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-3 > div.input-group > #selCh", "label=" + i.ToString());
                selenium.Type("css=#diConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-6 > div.input-group > #inpTag", "ABCDEFGHIJKLMNOPQRS" + (i + 30).ToString());
            }
            else
            {
                selenium.Select("id=selCh", "label=" + i.ToString());
                selenium.Type("id=inpTag", "ABCDEFGHIJKLMNOPQRS" + (i + 30).ToString());
            }
            selenium.Select("id=selMd", "label=High to Low Latch");
            selenium.Click("id=inpInv");
            if (dev.ModuleType.ToUpper() == "WISE-4012E")
                selenium.Click("css=#diConfigBaseForm > div.panel-footer.clearfix > div.pull-right > #btnSubmit");
            else
                selenium.Click("id=btnSubmit");
            System.Threading.Thread.Sleep(1000); PrintStep();
        }
        //for all
        if (dev.ModuleType.ToUpper() == "WISE-4012E")
            selenium.Select("css=#diConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-3 > div.input-group > #selCh", "label=All");
        else
            selenium.Select("id=selCh", "label=All");
        selenium.Select("id=selMd", "label=High to Low Latch");
        selenium.Click("id=inpInv");
        if (dev.ModuleType.ToUpper() == "WISE-4012E")
            selenium.Click("css=#diConfigBaseForm > div.panel-footer.clearfix > div.pull-right > #btnSubmit");
        else
            selenium.Click("id=btnSubmit");
        System.Threading.Thread.Sleep(1000);
        //
        PrintStep();

    }

    void DIExeFreqconfig()
    {
        PrintTitle("DIExeFreqconfig");
        selenium.Click("id=ioStatus0");
        if (dev.ModuleType.ToUpper() == "WISE-4012E") selenium.Click("link=DI");
        selenium.Click("id=diConfigRow0");
        if (dev.ModuleType.ToUpper() == "WISE-4012E")
            selenium.Click("xpath=(//a[contains(text(),'Configuration')])[2]");
        else selenium.Click("//a[contains(text(),'Configuration')]");
        PrintStep();
        //
        for (int i = 0; i < di_num; i++)
        {
            if (dev.ModuleType.ToUpper() == "WISE-4012E")
            {
                selenium.Select("css=#diConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-3 > div.input-group > #selCh", "label=" + i.ToString());
                selenium.Type("css=#diConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-6 > div.input-group > #inpTag", "ABCDEFGHIJKLMNOPQRS" + (i + 40).ToString());
            }
            else
            {
                selenium.Select("id=selCh", "label=" + i.ToString());
                selenium.Type("id=inpTag", "ABCDEFGHIJKLMNOPQRS" + (i + 40).ToString());
            }
            selenium.Select("id=selMd", "label=Frequency");
            selenium.Type("id=inpFqT", "255");
            if (dev.ModuleType.ToUpper() == "WISE-4012E")
                selenium.Click("css=#diConfigBaseForm > div.panel-footer.clearfix > div.pull-right > #btnSubmit");
            else
                selenium.Click("id=btnSubmit");
            System.Threading.Thread.Sleep(1000); PrintStep();
        }
        //for all
        if (dev.ModuleType.ToUpper() == "WISE-4012E")
            selenium.Select("css=#diConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-3 > div.input-group > #selCh", "label=All");
        else
            selenium.Select("id=selCh", "label=All");
        selenium.Select("id=selMd", "label=Frequency");
        selenium.Type("id=inpFqT", "255");
        if (dev.ModuleType.ToUpper() == "WISE-4012E")
            selenium.Click("css=#diConfigBaseForm > div.panel-footer.clearfix > div.pull-right > #btnSubmit");
        else
            selenium.Click("id=btnSubmit");
        System.Threading.Thread.Sleep(1000);
        //
        PrintStep();

    }

    void UDIExeDIconfig()
    {
        PrintTitle("UDIConfgSts");
        selenium.Click("id=ioStatus0");
        selenium.Click("link=DI");
        selenium.Click("id=diConfigRow0");
        selenium.Click("xpath=(//a[contains(text(),'Configuration')])[2]");
        PrintStep();
        //
        for (int i = 0; i < di_num; i++)
        {
            selenium.Select("css=#diConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-3 > div.input-group > #selCh", "label=" + i.ToString());
            selenium.Type("css=#diConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-6 > div.input-group > #inpTag", "ABCDEFGHIJKLMNOPQRST" + i.ToString());
            selenium.Select("id=selMd", "label=DI");
            selenium.Click("id=inpInv");
            selenium.Click("id=inpFltr");
            selenium.Type("id=inpFtLo", "65535");
            selenium.Type("id=inpFtHi", "65535");
            selenium.Click("css=#diConfigBaseForm > div.panel-footer.clearfix > div.pull-right > #btnSubmit");
            System.Threading.Thread.Sleep(1000); PrintStep();
        }
        //for all
        selenium.Select("css=#diConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-3 > div.input-group > #selCh", "label=All");
        selenium.Select("id=selMd", "label=DI");
        selenium.Click("id=inpInv");
        selenium.Click("id=inpFltr");
        selenium.Type("id=inpFtLo", "65535");
        selenium.Type("id=inpFtHi", "65535");
        selenium.Click("css=#diConfigBaseForm > div.panel-footer.clearfix > div.pull-right > #btnSubmit");
        System.Threading.Thread.Sleep(1000);
        //
        PrintStep();

    }

    void UDIExeCntconfig()
    {
        PrintTitle("DIExeCntconfig");
        selenium.Click("id=ioStatus0");
        selenium.Click("link=DI");
        selenium.Click("id=diConfigRow0");
        selenium.Click("xpath=(//a[contains(text(),'Configuration')])[2]");
        PrintStep();
        //
        for (int i = 0; i < di_num; i++)
        {
            selenium.Select("css=#diConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-3 > div.input-group > #selCh", "label=" + i.ToString());
            selenium.Type("css=#diConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-6 > div.input-group > #inpTag", "ABCDEFGHIJKLMNOPQRS" + (i * 10).ToString());
            selenium.Select("id=selMd", "label=Counter");
            selenium.Click("id=inpInv");
            selenium.Click("id=inpFltr");
            selenium.Type("id=inpFtLo", "65535");
            selenium.Type("id=inpFtHi", "65535");
            selenium.Type("id=inpCntIV", "4294967295");
            selenium.Click("id=inpCntKp");
            selenium.Click("css=#diConfigBaseForm > div.panel-footer.clearfix > div.pull-right > #btnSubmit");
            System.Threading.Thread.Sleep(1000); PrintStep();
        }
        //for all
        selenium.Select("css=#diConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-3 > div.input-group > #selCh", "label=All");
        selenium.Select("id=selMd", "label=Counter");
        selenium.Click("id=inpInv");
        selenium.Click("id=inpFltr");
        selenium.Type("id=inpFtLo", "65535");
        selenium.Type("id=inpFtHi", "65535");
        selenium.Type("id=inpCntIV", "4294967295");
        selenium.Click("id=inpCntKp");
        selenium.Click("css=#diConfigBaseForm > div.panel-footer.clearfix > div.pull-right > #btnSubmit");
        System.Threading.Thread.Sleep(1000);
        //
        PrintStep();
        //
        PrintTitle("Check DI Counter View");
        selenium.Click("xpath=(//a[contains(text(),'Status')])[3]");

        for (int i = 0; i < di_num; i++)
        {
            selenium.Click("id=switchonOffSwitchDI_" + i.ToString());
            selenium.Click("id=btnResetDI_" + i.ToString());
            selenium.Click("id=switchonOffSwitchDI_" + i.ToString());
            PrintStep();
        }

    }

    void UDIExeL2Hconfig()
    {
        PrintTitle("DIExeL2Hconfig");
        selenium.Click("id=ioStatus0");
        selenium.Click("link=DI");
        selenium.Click("id=diConfigRow0");
        selenium.Click("xpath=(//a[contains(text(),'Configuration')])[2]");
        PrintStep();
        //
        for (int i = 0; i < di_num; i++)
        {
            selenium.Select("css=#diConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-3 > div.input-group > #selCh", "label=" + i.ToString());
            selenium.Type("css=#diConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-6 > div.input-group > #inpTag", "ABCDEFGHIJKLMNOPQRS" + (i * 20).ToString());
            selenium.Select("id=selMd", "label=Low to High Latch");
            selenium.Click("id=inpInv");
            selenium.Click("css=#diConfigBaseForm > div.panel-footer.clearfix > div.pull-right > #btnSubmit");
            System.Threading.Thread.Sleep(1000); PrintStep();
        }
        //for all
        selenium.Select("css=#diConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-3 > div.input-group > #selCh", "label=All");
        selenium.Select("id=selMd", "label=Low to High Latch");
        selenium.Click("id=inpInv");
        selenium.Click("css=#diConfigBaseForm > div.panel-footer.clearfix > div.pull-right > #btnSubmit");
        System.Threading.Thread.Sleep(1000);
        //
        PrintStep();

    }

    void UDIExeH2Lconfig()
    {
        PrintTitle("DIExeH2Lconfig");
        selenium.Click("id=ioStatus0");
        selenium.Click("link=DI");
        selenium.Click("id=diConfigRow0");
        selenium.Click("xpath=(//a[contains(text(),'Configuration')])[2]");
        PrintStep();
        //
        for (int i = 0; i < di_num; i++)
        {
            selenium.Select("css=#diConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-3 > div.input-group > #selCh", "label=" + i.ToString());
            selenium.Type("css=#diConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-6 > div.input-group > #inpTag", "ABCDEFGHIJKLMNOPQRS" + (i * 30).ToString());
            selenium.Select("id=selMd", "label=High to Low Latch");
            selenium.Click("id=inpInv");
            selenium.Click("css=#diConfigBaseForm > div.panel-footer.clearfix > div.pull-right > #btnSubmit");
            System.Threading.Thread.Sleep(1000); PrintStep();
        }
        //for all
        selenium.Select("css=#diConfigBaseForm > div.panel-heading > div.form-group > div.col-lg-3 > div.input-group > #selCh", "label=All");
        selenium.Select("id=selMd", "label=High to Low Latch");
        selenium.Click("id=inpInv");
        selenium.Click("css=#diConfigBaseForm > div.panel-footer.clearfix > div.pull-right > #btnSubmit");
        System.Threading.Thread.Sleep(1000);
        //
        PrintStep();

    }

}//class
