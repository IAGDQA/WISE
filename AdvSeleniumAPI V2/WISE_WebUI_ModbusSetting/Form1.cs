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
        PrintTitle("Modubus Coils");
        selenium.Click("id=configuration");
        selenium.Click("link=Modbus");
        if(dev.ModuleType == "WISE-4050" || dev.ModuleType == "WISE-4060")
        {
            selenium.Type("id=basDI", "11");
            selenium.Type("id=basDO", "171");
            selenium.Type("id=basCtS", "331");
            selenium.Type("id=basCtClr", "371");
            selenium.Type("id=basCtOv", "411");
            selenium.Type("id=basLch", "451");
            selenium.Type("id=basLB", "9999");
            selenium.Click("id=btnModbusCoilSubmit");
        }
        else if(dev.ModuleType == "WISE-4012")
        {
            selenium.Type("id=basDI", "11");
            selenium.Type("id=basDO", "171");
            selenium.Type("id=basCtS", "331");
            selenium.Type("id=basCtClr", "37");
            selenium.Type("id=basCtClr", "371");
            selenium.Type("id=basCtOv", "411");
            selenium.Type("id=basLch", "451");
            selenium.Type("id=basAIHR", "1011");
            selenium.Type("id=basAILR", "1111");
            selenium.Type("id=basAIB", "1211");
            selenium.Type("id=basHAlm", "1311");
            selenium.Type("id=basLAlm", "1411");
            selenium.Type("id=basLB", "9999");
            selenium.Click("id=btnModbusCoilSubmit");
        }
        else if (dev.ModuleType == "WISE-4012")
        {
            selenium.Type("id=basLB", "9999");
            selenium.Type("id=basExB", "8888");
            selenium.Type("id=basLch", "571");
            selenium.Type("id=basCtOv", "491");
            selenium.Type("id=basCtClr", "411");
            selenium.Type("id=basCtS", "331");
            selenium.Type("id=basDI", "11");
            selenium.Click("id=btnModbusCoilSubmit");
        }

        PrintStep();
        PrintTitle("Modubus Registors");
        selenium.Click("id=modbusAddrRegConfig");
        if (dev.ModuleType == "WISE-4050" || dev.ModuleType == "WISE-4060")
        {
            selenium.Type("id=basPsLo", "91");
            selenium.Type("id=basPsHi", "171");
            selenium.Type("id=basCtFq", "11");         
            selenium.Type("id=basPsAV", "251");
            selenium.Type("id=basPsIV", "331");
            selenium.Type("id=basMNm", "2111");
            selenium.Type("xpath=(//input[@id='basDI'])[2]", "3011");
            selenium.Type("xpath=(//input[@id='basDO'])[2]", "3031");
            selenium.Type("id=basCntIV", "4011");
        }
        else if(dev.ModuleType == "WISE-4012")
        {
            selenium.Type("id=basCntIV", "4011");
            selenium.Type("xpath=(//input[@id='basDO'])[2]", "3031");
            selenium.Type("xpath=(//input[@id='basDI'])[2]", "3011");
            selenium.Type("id=basAICh", "2211");
            selenium.Type("id=basMNm", "2111");
            selenium.Type("id=basAIPF", "2311");
            selenium.Type("id=basAICd", "2011");
            selenium.Type("id=basAISc", "1911");
            selenium.Type("id=basHisLF", "1711");
            selenium.Type("id=basHisHF", "1511");
            selenium.Type("id=basAIF", "1311");
            selenium.Type("id=basHisL", "1211");
            selenium.Type("id=basHisH", "1111");
            selenium.Type("id=basAIFl", "1011");
            selenium.Type("id=basPsIV", "371");
            selenium.Type("id=basPsAV", "331");
            selenium.Type("id=basPsHi", "291");
            selenium.Type("id=basPsLo", "251");
            selenium.Type("id=basCtFq", "171");
            selenium.Type("id=basAI", "11");
        }
        else if (dev.ModuleType == "WISE-4051")
        {
            selenium.Type("id=basLg", "9999");
            selenium.Type("id=basExWE", "8888");
            selenium.Type("id=basExBE", "7777");
            selenium.Type("id=basExW", "6666");
            selenium.Type("xpath=(//input[@id='basDI'])[2]", "3011");
            selenium.Type("id=basMNm", "2111");
            selenium.Type("id=basCtFq", "11");
        }
        selenium.Type("id=basLg", "9999");
        //Only wireless module support.
        if (dev.ModuleType != "WISE-4050/LAN"
            || dev.ModuleType != "WISE-4060/LAN"
            || dev.ModuleType != "WISE-4010/LAN")
        {
            selenium.Type("id=basRssi", "9998");
        }
            
        selenium.Click("id=btnModbusRegSubmit");

        PrintStep();
        selenium.Close();
    }
}//class
