using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using AdvSeleniumAPI;
using iATester;
using Model;
using Service;
using System.IO;
using System.Web.Script.Serialization;
using System.Threading;
using System.Text.RegularExpressions;

public partial class WISE_PrivateServer : Form, iATester.iCom
{
    private delegate void DataGridViewCtrlAddDataRow(DataGridViewRow i_Row);
    private DataGridViewCtrlAddDataRow m_DataGridViewCtrlAddDataRow;
    internal const int Max_Rows_Val = 65535;

    IHttpReqService HttpReqService;
    DataHandleService dataHld;
    DeviceModel dev;
    string AddressIP = "", devName = "", path = "", browser = "";
    bool ConnectFlg = false, ResFlg = false, Listening = false;
    int errorCntStep = 0;
    string devMac = ""; string DUT_mod = "";
    AdvSeleniumAPIv2 selenium; IniDataFmt LoadPara;

    //
    internal const string HTTP_Prefix = "http://*:8000/"; //Waiting for HTTP request on port 8000
    //internal const string HTTPS_Prefix = "https://*:1234/"; //Waiting for HTTPS request on port 8080
    internal const string Url_File_UploadLog_Token = @"upload_log";
    internal const string Url_Json_IoLog_Token = @"io_log";
    internal const string Url_Json_SysLog_Token = @"sys_log";
    internal const string DataType_Csv_File = @"CSV File";
    internal const string DataType_Json_Str = @"JSON String";
    internal const string Slash_Str_Token = @"/";
    internal const string BackSlash_Str_Token = @"\";
    static HttpListener m_HttpListener;
    private Thread httpRequestThread;
    private long _runState = (long)State.Stopped;

    //iATester
    //Send Log data to iAtester
    public event EventHandler<LogEventArgs> eLog = delegate { };
    //Send test result to iAtester
    public event EventHandler<ResultEventArgs> eResult = delegate { };
    //Send execution status to iAtester
    public event EventHandler<StatusEventArgs> eStatus = delegate { };

    //
    public WISE_PrivateServer()
    {
        InitializeComponent();
    }

    private void WISE_PrivateServer_Load(object sender, EventArgs e)
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
        //
        //WorkSteps();
    }
    private void LoadFile()
    {
        LoadPara = dataHld.GetPara(Application.StartupPath);
        textBox1.Text = AddressIP = LoadPara.IP;
    }
    private void WISE_PrivateServer_FormClosing(object sender, FormClosingEventArgs e)
    {
        StopListen();
    }
    public void StartTest()//iATester
    {
        LoadFile();
        if (ExeConnectionDUT())
        {
            eStatus(this, new StatusEventArgs(iStatus.Running));
            WorkSteps();
            //
            if (!ResFlg || errorCntStep > 0)
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

    public enum State
    {
        Stopped,
        Stopping,
        Starting,
        Started
    }

    public State RunState
    {
        get
        {
            return (State)Interlocked.Read(ref _runState);
        }
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
                errorCntStep++;
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
        ResFlg = false; errorCntStep = 0;
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
        if(ipList.Count < 1)
        {
            PrintTitle("Get Host IP fail.");
            return;
        }
        string[] HostIP = new string[ipList.Count];
        int i = 0;
        foreach (var item in ipList)
        {
            HostIP[i] = (string)item;
            i++;
        }
        //
//#if Debug
        devMac = "";
        selenium = new AdvSeleniumAPIv2();

        selenium.StartupServer("http://" + textBox1.Text);
        System.Threading.Thread.Sleep(1000);

        selenium.Type("id=ACT0", "root");
        selenium.Type("id=PWD0", "00000000");
        selenium.Click("id=APY0");
        selenium.WaitForPageToLoad("30000");
        PrintStep();
        //先確認DUT在哪個MODE底下
        PrintTitle("Check DUT mode");
        var mod = DUT_mod = selenium.GetValue("id=inpWMd");
        PrintStep();
        //
        if (mod == "Normal Mode")
        {
            PrintTitle("Get device MAC address");
            selenium.Click("id=configuration");
            if (dev.ModuleType.ToUpper() == "WISE-4050/LAN"
                || dev.ModuleType.ToUpper() == "WISE-4060/LAN"
                || dev.ModuleType.ToUpper() == "WISE-4010/LAN")
            {
                selenium.Click("link=Network");
            }
            else
            {
                selenium.Click("link=Wireless");
            }
                
            devMac = selenium.GetValue("id=inpMAC");
            PrintStep();
            PrintTitle("Get MAC address is [" + devMac + "]");
            //
            PrintTitle("Check Private Server items");
            selenium.Click("id=configuration");
            selenium.Click("link=Cloud");
            selenium.Select("id=selCloud", "label=  Private Server");
            selenium.Type("id=logIP", HostIP[0]);
            selenium.Type("id=logPWeb", "8000");
            selenium.Type("id=logUurl", "/upload_log");
            selenium.Type("id=logDurl", "/io_log");
            selenium.Type("id=logSurl", "/sys_log");
            selenium.Click("id=RadioSslDisable");
            selenium.Click("id=btnPrivateServerSubmit");
            selenium.Type("id=logPu", "root");
            selenium.Type("id=logPw", "00000000");
            PrintStep();
            //
            PrintTitle("Enable [By Period] checkbox");
            selenium.Click("id=configuration");
            selenium.Click("id=advancedFunction");
            selenium.Click("id=dataLog");
            selenium.Type("id=inpPItv", "1");
            var res = selenium.GetValue("id=inpPer");
            if (selenium.GetValue("id=inpPer") == "off")
                selenium.Click("id=inpPer");
            selenium.Click("id=btnLogConfigAll");
            PrintStep();
            //
            selenium.Click("id=advancedFunction");
            selenium.Click("id=dataLog");
            selenium.Click("link=Logger Configuration");
            PrintTitle("Enable [IO Log] checkbox");
            if (selenium.GetValue("id=memDEn") == "off")
                selenium.Click("id=memDEn");
            PrintStep();
            PrintTitle("Enable [Cloud Upload] checkbox");
            if (selenium.GetValue("id=cloudEn") == "off")
                selenium.Click("id=cloudEn");
            selenium.Select("id=selDEn", "label=Item Periodic Interval mode");
            selenium.Type("id=inpDItm", "1");
            selenium.Type("id=inpDTag", "WISE_PrvSrv_AUTOTEST");
            if (selenium.GetValue("id=pushDEn") == "off")
                selenium.Click("css=#push_output > div.form-group.row > div.col-lg-12 > div.col-sm-10 > div.input-group > div.SliderSwitch > label.SliderSwitch-label > span.SliderSwitch-inner");
            selenium.Click("id=btnLoggerConfigSubmit");
            PrintStep();
        }
        else PrintTitle("Mode is not in [Normal Mode].");
        //
        selenium.Close();
//#endif
        //
        StartListener();
        int WDT = 0;
        while (Listening)
        {
            PrintTitle("Listening");
            Application.DoEvents();
            if (WDT > 999)
            {
                PrintTitle("Timeout....");
                break;
            }
            WDT++;
        }
    }

    void StartListener()
    {
        counter = 0; errorCnt = 0;
        try
        {
            m_HttpListener = new HttpListener();

            try
            {
                m_HttpListener.Prefixes.Add(HTTP_Prefix);
                //m_HttpListener.Prefixes.Add(HTTPS_Prefix);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Err : " + ex.ToString());
            }
            try
            {
                m_HttpListener.AuthenticationSchemes = AuthenticationSchemes.Anonymous;
                httpRequestThread = new Thread(new ThreadStart(StartListen));
                httpRequestThread.IsBackground = true;
                httpRequestThread.Start();
                PrintTitle("httpRequestThread.Start()");
            }
            catch (Exception exe)
            {
                MessageBox.Show("Err : " + exe.ToString());
            }
        }
        catch (Exception eee)
        {
            MessageBox.Show("Err : " + eee.ToString());
        }
    }
    private void StartListen()
    {
        Listening = true;
        Interlocked.Exchange(ref this._runState, (long)State.Starting);
        try
        {
            if (!m_HttpListener.IsListening)
            {
                try
                {
                    m_HttpListener.Start();
                    PrintTitle("m_HttpListener.Start()");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error:\n TCP port listen failed.\n Please make sure:\n(1) Port is not used.\n(2) Firewall is not blocked.\n(3) Use Administrator to run this program.\n\nDetail Information:\n" + ex.ToString());
                }
            }
            if (m_HttpListener.IsListening)
            {
                Interlocked.Exchange(ref this._runState, (long)State.Started);
            }
            try
            {
                while (RunState == State.Started)
                {
                    HttpListenerContext context = m_HttpListener.GetContext();
                    IncomingHttpRequest(context);
                }
            }
            catch (HttpListenerException)
            {
                Interlocked.Exchange(ref this._runState, (long)State.Stopped);
            }
        }
        finally
        {
            Interlocked.Exchange(ref this._runState, (long)State.Stopped);
        }
        //
        //if (counter > 10) StopListen();
    }

    private void StopListen()
    {
        if (m_HttpListener == null) return;
        if (m_HttpListener.IsListening)
        {
            PrintTitle("StopListen");
            m_HttpListener.Stop();
            Interlocked.Exchange(ref this._runState, (long)State.Stopping);
            Listening = false;
        }

        //httpRequestThread.Join();
        if (httpRequestThread != null && httpRequestThread.IsAlive)
        {
            httpRequestThread.Join(1000);//block main thread and wait max 2 seconds for http thread terminate
            httpRequestThread.Abort();
        }
    }

    private string getUniqueFileName(string logFilePath)
    {
        while (true)
        {
            if (File.Exists(logFilePath))
            {
                //file already exists, create new file name
                string extensionName = Path.GetExtension(logFilePath);
                var fileName = Path.GetFileNameWithoutExtension(logFilePath);
                var dirName = Path.GetDirectoryName(logFilePath);
                if (fileName.IndexOf(' ') == -1)
                {
                    //ex: 20151125111821
                    fileName = fileName + " (1)" + extensionName;
                }
                else
                {
                    //ex: 20151125111821 (1)
                    string pattern = @"(\S+)\ \((\d+)\)";
                    // Instantiate the regular expression object.
                    Regex r = new Regex(pattern, RegexOptions.IgnoreCase);
                    // Match the regular expression pattern against a text string.
                    Match m = r.Match(fileName);
                    if (m.Success)
                    {
                        string match1 = m.Groups[1].ToString();
                        string match2 = m.Groups[2].ToString();
                        int tmp = Int32.Parse(match2) + 1;
                        fileName = match1 + " (" + tmp + ")" + extensionName;
                    }
                    else
                    {
                        //match failed, append a random number to file name
                        Random rnd = new Random();
                        int tmp = rnd.Next(1, 1000);
                        fileName = fileName + "_" + tmp + extensionName;
                    }
                }
                logFilePath = Path.Combine(dirName, fileName);
            }
            else
            {
                //file not exists
                return logFilePath;
            }
        }
    }
    int counter = 0, errorCnt = 0;
    private void IncomingHttpRequest(HttpListenerContext context)
    {
        PrintTitle("IncomingHttpRequest");
        HttpListenerResponse response = context.Response;
        var data_text = new StreamReader(context.Request.InputStream, context.Request.ContentEncoding).ReadToEnd();
        var receive_data = System.Web.HttpUtility.UrlDecode(data_text);
        try {
            JavaScriptSerializer serializer = new JavaScriptSerializer();
            var values = serializer.Deserialize<SysData>(receive_data);
            PrintTitle("Get incoming Message MAC is [" + values.MAC + "]");
            PrintTitle(receive_data);
            if (values.MAC == devMac)
            {
                PrintTitle("Device MAC is correct.");
                ResFlg = true;
            }                
            else
            { PrintTitle("MAC verified fail."); errorCnt++; }
        }
        catch(Exception ex)
        {
            MessageBox.Show(ex.ToString());
        }        

        if (context.Request.Url.AbsolutePath.Contains(Url_File_UploadLog_Token))
        {
            string szFileName = string.Empty;
            List<string> szList = new List<string>(context.Request.Url.AbsolutePath.Split('/'));
            foreach (string subValue in szList)
            {
                if (subValue.Contains(".csv"))
                {
                    szFileName = context.Request.Url.AbsolutePath.Replace(Slash_Str_Token, BackSlash_Str_Token);
                    szFileName = szFileName.Substring(BackSlash_Str_Token.Length, (szFileName.Length - BackSlash_Str_Token.Length));
                    szFileName = getUniqueFileName(szFileName);
                    FileInfo file = new FileInfo(szFileName);
                    file.Directory.Create();
                    System.IO.File.WriteAllText(file.FullName, receive_data);
                    break;
                }
            }
        }
        response.StatusCode = 200;
        response.StatusDescription = "OK";
        response.Close();
        
        if (counter > 10) StopListen();
        else counter++;
    }

    public class SysData
    {
        public int PE { get; set; }
        public string TIM { get; set; }
        public string UID { get; set; }
        public string MAC { get; set; }
        public object Record { get; set; }
    }
}//class
