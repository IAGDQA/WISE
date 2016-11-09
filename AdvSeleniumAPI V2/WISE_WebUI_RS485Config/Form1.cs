using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
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
        errorCnt = 0;
        if (dev.ModuleType == "WISE-4051")
        {
            selenium = new AdvSeleniumAPIv2();

            selenium.StartupServer("http://" + textBox1.Text);
            System.Threading.Thread.Sleep(1000);

            selenium.Type("id=ACT0", "root");
            selenium.Type("id=PWD0", "00000000");
            selenium.Click("id=APY0");
            selenium.WaitForPageToLoad("30000");
            PrintStep();
            for (int i = 0; i < 1; i++)
            {
                selenium.Click("id=ioStatus0");
                selenium.Click("link=COM1");
                selenium.Click("link=Modbus/RTU Configuration");
                if (i > 0)
                {
                    selenium.Select("//div[@id='TabCom1']/div/div/div/div[2]/div/div[2]/div/div/form/div/div/div/div/select"
                        , "label=9600 bps");
                    selenium.Select("//div[@id='TabCom1']/div/div/div/div[2]/div/div[2]/div/div/form/div/div/div[2]/div/select"
                        , "label=7 bit");
                    selenium.Select("//div[@id='TabCom1']/div/div/div/div[2]/div/div[2]/div/div/form/div/div/div[3]/div/select"
                        , "label=Odd");
                    selenium.Select("//div[@id='TabCom1']/div/div/div/div[2]/div/div[2]/div/div/form/div/div/div[4]/div/select"
                        , "label=1 bit");
                    PrintStep();
                }
                else
                {
                    selenium.Select("//div[@id='TabCom1']/div/div/div/div[2]/div/div[2]/div/div/form/div/div/div/div/select"
                        , "label=115200 bps");
                    selenium.Select("//div[@id='TabCom1']/div/div/div/div[2]/div/div[2]/div/div/form/div/div/div[2]/div/select"
                        , "label=8 bit");
                    selenium.Select("//div[@id='TabCom1']/div/div/div/div[2]/div/div[2]/div/div/form/div/div/div[3]/div/select"
                        , "label=Even");
                    selenium.Select("//div[@id='TabCom1']/div/div/div/div[2]/div/div[2]/div/div/form/div/div/div[4]/div/select"
                        , "label=2 bit");
                    PrintStep();
                }

                selenium.Type("xpath=(//input[@type='number'])[11]", "5000");
                selenium.Type("xpath=(//input[@type='number'])[12]", "1000");
                selenium.Click("name=004");
                //
                selenium.Click("//div[@id='TabCom1']/div/div/div/div[2]/div/div[2]/div/div/form/div[2]/div/button");
                selenium.Click("link=Rule Setting");
                PrintStep();

                //rule 01
                selenium.Type("xpath=(//input[@type='number'])[13]", "20" + i.ToString());
                selenium.Select("//div[@id='TabCom1']/div/div/div/div[2]/div/div[2]/div/div[2]/form/div/table/tbody/tr/td[3]/select"
                    , "label=01 Coil status");
                selenium.Type("xpath=(//input[@type='number'])[14]", "999" + i.ToString());
                selenium.Type("xpath=(//input[@type='number'])[15]", "4");
                selenium.Select("//div[@id='TabCom1']/div/div/div/div[2]/div/div[2]/div/div[2]/form/div/table/tbody/tr/td[6]/select"
                    , "label=R/W");
                selenium.Type("xpath=(//input[@type='number'])[16]", "999" + i.ToString());
                selenium.Click("xpath=(//input[@type='checkbox'])[19]");
                PrintStep();

                //rule 02
                selenium.Type("xpath=(//input[@type='number'])[18]", "20" + (i + 1).ToString());
                selenium.Select("//div[@id='TabCom1']/div/div/div/div[2]/div/div[2]/div/div[2]/form/div/table/tbody/tr[2]/td[3]/select"
                    , "label=02 Input status");
                selenium.Type("xpath=(//input[@type='number'])[19]", "999" + i.ToString());
                selenium.Type("xpath=(//input[@type='number'])[20]", "4");
                selenium.Select("//div[@id='TabCom1']/div/div/div/div[2]/div/div[2]/div/div[2]/form/div/table/tbody/tr[2]/td[6]/select"
                    , "label=R/W");
                selenium.Type("xpath=(//input[@type='number'])[21]", "999" + i.ToString());
                selenium.Type("xpath=(//input[@type='number'])[22]", "4");
                selenium.Click("xpath=(//input[@type='checkbox'])[20]");
                PrintStep();

                //rule 03
                selenium.Type("xpath=(//input[@type='number'])[23]", "20" + (i + 2).ToString());
                selenium.Select("//div[@id='TabCom1']/div/div/div/div[2]/div/div[2]/div/div[2]/form/div/table/tbody/tr[3]/td[3]/select"
                    , "label=01 Coil status");
                selenium.Type("xpath=(//input[@type='number'])[24]", "999" + i.ToString());
                selenium.Type("xpath=(//input[@type='number'])[25]", "4");
                selenium.Select("//div[@id='TabCom1']/div/div/div/div[2]/div/div[2]/div/div[2]/form/div/table/tbody/tr[3]/td[6]/select"
                    , "label=R/W");
                selenium.Type("xpath=(//input[@type='number'])[26]", "999" + i.ToString());
                selenium.Type("xpath=(//input[@type='number'])[27]", "8");
                selenium.Click("xpath=(//input[@type='checkbox'])[21]");
                PrintStep();

                //rule 04
                selenium.Type("xpath=(//input[@type='number'])[28]", "20" + (i + 3).ToString());
                selenium.Select("//div[@id='TabCom1']/div/div/div/div[2]/div/div[2]/div/div[2]/form/div/table/tbody/tr[4]/td[3]/select"
                    , "label=02 Input status");
                selenium.Type("xpath=(//input[@type='number'])[29]", "999" + i.ToString());
                selenium.Type("xpath=(//input[@type='number'])[30]", "4");
                selenium.Select("//div[@id='TabCom1']/div/div/div/div[2]/div/div[2]/div/div[2]/form/div/table/tbody/tr[4]/td[6]/select"
                    , "label=R/W");
                selenium.Type("xpath=(//input[@type='number'])[31]", "999" + i.ToString());
                selenium.Type("xpath=(//input[@type='number'])[32]", "12");
                selenium.Click("xpath=(//input[@type='checkbox'])[22]");
                PrintStep();

                //rule 05
                selenium.Type("xpath=(//input[@type='number'])[33]", "20" + (i + 4).ToString());
                selenium.Select("//div[@id='TabCom1']/div/div/div/div[2]/div/div[2]/div/div[2]/form/div/table/tbody/tr[5]/td[3]/select"
                    , "label=03 Holding register");
                selenium.Type("xpath=(//input[@type='number'])[34]", "999" + i.ToString());
                selenium.Type("xpath=(//input[@type='number'])[35]", "4");
                selenium.Select("//div[@id='TabCom1']/div/div/div/div[2]/div/div[2]/div/div[2]/form/div/table/tbody/tr[5]/td[6]/select"
                    , "label=R/W");
                selenium.Type("xpath=(//input[@type='number'])[36]", "999" + i.ToString());
                selenium.Type("xpath=(//input[@type='number'])[37]", "16");
                selenium.Click("xpath=(//input[@type='checkbox'])[23]");
                PrintStep();

                //rule 06
                selenium.Type("xpath=(//input[@type='number'])[38]", "20" + (i + 5).ToString());
                selenium.Select("//div[@id='TabCom1']/div/div/div/div[2]/div/div[2]/div/div[2]/form/div/table/tbody/tr[6]/td[3]/select"
                    , "label=04 Input register");
                selenium.Type("xpath=(//input[@type='number'])[39]", "999" + i.ToString());
                selenium.Type("xpath=(//input[@type='number'])[40]", "4");
                selenium.Select("//div[@id='TabCom1']/div/div/div/div[2]/div/div[2]/div/div[2]/form/div/table/tbody/tr[6]/td[6]/select"
                    , "label=R/W");
                selenium.Type("xpath=(//input[@type='number'])[41]", "999" + i.ToString());
                selenium.Type("xpath=(//input[@type='number'])[42]", "20");
                selenium.Click("xpath=(//input[@type='checkbox'])[24]");
                PrintStep();

                //rule 07
                selenium.Type("xpath=(//input[@type='number'])[43]", "20" + (i + 6).ToString());
                selenium.Select("//div[@id='TabCom1']/div/div/div/div[2]/div/div[2]/div/div[2]/form/div/table/tbody/tr[7]/td[3]/select"
                    , "label=03 Holding register");
                selenium.Type("xpath=(//input[@type='number'])[44]", "999" + i.ToString());
                selenium.Type("xpath=(//input[@type='number'])[45]", "4");
                selenium.Select("//div[@id='TabCom1']/div/div/div/div[2]/div/div[2]/div/div[2]/form/div/table/tbody/tr[7]/td[6]/select"
                    , "label=R/W");
                selenium.Type("xpath=(//input[@type='number'])[46]", "999" + i.ToString());
                selenium.Type("xpath=(//input[@type='number'])[47]", "24");
                selenium.Click("xpath=(//input[@type='checkbox'])[25]");
                PrintStep();

                //rule 08
                selenium.Type("xpath=(//input[@type='number'])[48]", "20" + (i + 7).ToString());
                selenium.Select("//div[@id='TabCom1']/div/div/div/div[2]/div/div[2]/div/div[2]/form/div/table/tbody/tr[8]/td[3]/select"
                    , "label=04 Input register");
                selenium.Type("xpath=(//input[@type='number'])[49]", "999" + i.ToString());
                selenium.Type("xpath=(//input[@type='number'])[50]", "4");
                selenium.Select("//div[@id='TabCom1']/div/div/div/div[2]/div/div[2]/div/div[2]/form/div/table/tbody/tr[8]/td[6]/select"
                    , "label=R/W");
                selenium.Type("xpath=(//input[@type='number'])[51]", "999" + i.ToString());
                selenium.Type("xpath=(//input[@type='number'])[52]", "28");
                selenium.Click("xpath=(//input[@type='checkbox'])[25]");
                PrintStep();

                //
                selenium.Click("//div[@id='TabCom1']/div/div/div/div[2]/div/div[2]/div/div[2]/form/div[2]/button");
            }

            selenium.Close();
        }
        else
            PrintTitle("Module not support.");
    }
}//class
