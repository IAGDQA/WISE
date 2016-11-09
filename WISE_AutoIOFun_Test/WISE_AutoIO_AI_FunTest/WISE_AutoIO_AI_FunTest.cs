using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Threading;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Runtime.InteropServices;
using Model;
using Service;
using Messengers;
using System.Collections.ObjectModel;
using iATester;

namespace WISE_AutoIO_AI_FunTest
{
    public partial class WISE_AutoIO_AI_FunTest : Form, iATester.iCom
    {
        CheckBox[] chkbox = new CheckBox[9];
        TextBox[] setTxtbox = new TextBox[9];
        TextBox[] getTxtbox = new TextBox[9];
        TextBox[] apaxTxtbox = new TextBox[9];
        TextBox[] modbTxtbox = new TextBox[9];
        Label[] resLabel = new Label[9];

        private delegate void DataGridViewCtrlAddDataRow(DataGridViewRow i_Row);
        private DataGridViewCtrlAddDataRow m_DataGridViewCtrlAddDataRow;
        internal const int Max_Rows_Val = 65535;
        //
        //DispatcherTimer timer = new DispatcherTimer();
        bool ConnReadyFlg = false;//because display would get error.
        //string tempIP = "192.168.0.157";
        string tempIP = "192.168.1.1";
        DeviceModel Device;
        DeviceModel APAX5070;
        WISE_IO_Model Ref_IO_Mod;
        string filename = "WISE_IO_CONFIG.ini";
        string folderPath = "";
        //
        private IHttpReqService HttpReqService;
        private IModbusTCPService ModbusTCPService;//20151016
        private IAPAX5070Service APAX5070Service;//20160316
        //
        int mainTaskStp = 0, FailCnt = 0;
        int ai_id_offset = 40;
        int do_id_offset = 20;
        int DelayT = 500;
        int[] temp_rng; //ai range change used
        int VModNum, AModNum;//Voltage/Current Mode number
        object obj;
        int RngMod = 323;//20161101 fix wise4010 problem(current mode)
        //
        //iATester
        //Send Log data to iAtester
        public event EventHandler<LogEventArgs> eLog = delegate { };
        //Send test result to iAtester
        public event EventHandler<ResultEventArgs> eResult = delegate { };
        //Send execution status to iAtester
        public event EventHandler<StatusEventArgs> eStatus = delegate { };
        //========================================================//
        #region Observable Properties
        int _reftypIdx = 0;
        public int RefTypSelIdx
        {
            get { return _reftypIdx; }
            set
            {
                _reftypIdx = value;
                //RefreshChTypeItem();
                //RaisePropertyChanged("RefTypSelIdx");
            }

        }
        int _refchtypIdx = 0;
        public int ChSelIdx
        {
            get { return _refchtypIdx; }
            set
            {
                _refchtypIdx = value;
                //RaisePropertyChanged("ChSelIdx");
            }

        }
        int _refvatypIdx = 0;
        public int VASelIdx
        {
            get { return _refvatypIdx; }
            set
            {
                _refvatypIdx = value;
                //RaisePropertyChanged("VASelIdx");
            }

        }
        int _refDevtypIdx = 0;
        public int CtrlDevSelIdx
        {
            get { return _refDevtypIdx; }
            set
            {
                _refDevtypIdx = value;
                //RaisePropertyChanged("CtrlDevSelIdx");
            }

        }
        int _apaxAOslot_idx = 1;//for modbus/apax5580
        public int OutAO_Slot_Idx
        {
            get { return _apaxAOslot_idx; }
            set
            {
                _apaxAOslot_idx = value;
                idxTxtbox.Text = _apaxAOslot_idx.ToString();
                //RaisePropertyChanged("OutAO_Slot_Idx");
            }

        }
        int _ref_ai_num = 0;
        public string OutAO_Ch_Len
        {
            get
            { return "0 ~ " + (_ref_ai_num - 1).ToString(); }
            set
            {
                //RaisePropertyChanged("OutAO_Ch_Len");
            }
        }
        public ObservableCollection<string> ModelType
        {
            get;
            set;
        }
        ObservableCollection<string> _chTyp = new ObservableCollection<string>();
        public ObservableCollection<string> ChannelType
        {
            get
            {
                return _chTyp;
            }
            set
            {
                _chTyp = value;
                UpdateCh_Typ(_chTyp);
            }
        }
        ObservableCollection<string> _vaTyp = new ObservableCollection<string>();
        public ObservableCollection<string> VAType
        {
            get
            {
                return _vaTyp;
            }
            set
            {
                _vaTyp = value;
                UpdateVATyp(_vaTyp);
            }
        }
        //public ObservableCollection<ListViewData> ProcessView
        //{
        //    get;
        //    set;
        //}
        string _runStr = "Run";
        public string RunBtnStr
        {
            get { return _runStr; }
            set
            {
                _runStr = value;
                //RaisePropertyChanged("RunBtnStr");
            }

        }
        string _testRes = "N/A";
        public string TestResult
        {
            get { return _testRes; }
            set
            {
                _testRes = value;
                labelRes.Text = _testRes;
                //RaisePropertyChanged("TestResult");
            }

        }
        //
        int[] setData = new int[10];
        public int SetData01
        {
            get { return setData[1]; }
            set
            {
                setTxtbox[0].Text = value.ToString();
                setData[1] = value;
            }
        }
        public int SetData02
        {
            get { return setData[2]; }
            set
            {
                setTxtbox[1].Text = value.ToString();
                setData[2] = value;
            }
        }
        public int SetData03
        {
            get { return setData[3]; }
            set
            {
                setTxtbox[2].Text = value.ToString();
                setData[3] = value;
            }
        }
        string setData04 = "";
        public string SetData04
        {
            get { return setData04; }
            set
            {
                setTxtbox[3].Text = value;
                setData04 = value;
            }
        }
        string setData05 = "";
        public string SetData05
        {
            get { return setData05; }
            set
            {
                setTxtbox[5].Text = value;
                setData05 = value;
            }
        }
        string setData06 = "";
        public string SetData06
        {
            get { return setData06; }
            set
            {
                setTxtbox[6].Text = value;
                setData06 = value;
            }
        }
        string setData07 = "";
        public string SetData07
        {
            get { return setData07; }
            set
            {
                setTxtbox[8].Text = value;
                setData07 = value;
            }
        }
        public int SetData08
        {
            get { return setData[8]; }
            set
            {
                setTxtbox[4].Text = value.ToString();
                setData[8] = value;
            }
        }
        public int SetData09
        {
            get { return setData[9]; }
            set
            {
                setTxtbox[7].Text = value.ToString();
                setData[9] = value;
            }
        }
        //
        string[] resStr = new string[10];
        public string ProcessStep
        {
            get { return resStr[0]; }
            set
            {
                resStr[0] = value;
                tssLabel2.Text = resStr[0];
            }
        }
        public string ResStr01
        {
            get { return resStr[1]; }
            set
            {
                resLabel[0].Text = value.ToString();
                resStr[1] = value;
            }
        }
        public string ResStr02
        {
            get { return resStr[2]; }
            set
            {
                resLabel[1].Text = value.ToString();
                resStr[2] = value;
            }
        }
        public string ResStr03
        {
            get { return resStr[3]; }
            set
            {
                resLabel[2].Text = value.ToString();
                resStr[3] = value;
            }
        }
        public string ResStr04
        {
            get { return resStr[4]; }
            set
            {
                resLabel[3].Text = value.ToString();
                resStr[4] = value;
            }
        }
        public string ResStr05
        {
            get { return resStr[5]; }
            set
            {
                resLabel[5].Text = value.ToString();
                resStr[5] = value;
            }
        }
        public string ResStr06
        {
            get { return resStr[6]; }
            set
            {
                resLabel[6].Text = value.ToString();
                resStr[6] = value;
            }
        }
        public string ResStr07
        {
            get { return resStr[7]; }
            set
            {
                resLabel[8].Text = value.ToString();
                resStr[7] = value;
            }
        }
        public string ResStr08
        {
            get { return resStr[8]; }
            set
            {
                resLabel[4].Text = value.ToString();
                resStr[8] = value;
            }
        }
        public string ResStr09
        {
            get { return resStr[9]; }
            set
            {
                resLabel[7].Text = value.ToString();
                resStr[9] = value;
            }
        }
        //
        int[] outData = new int[10];
        public int OutData01
        {
            get { return outData[1]; }
            set
            {
                apaxTxtbox[0].Text = value.ToString();
                outData[1] = value;
            }
        }

        public int OutData02
        {
            get { return outData[2]; }
            set
            {
                apaxTxtbox[1].Text = value.ToString();
                outData[2] = value;
            }
        }

        public int OutData03
        {
            get { return outData[3]; }
            set
            {
                apaxTxtbox[2].Text = value.ToString();
                outData[3] = value;
            }
        }

        public int OutData04
        {
            get { return outData[4]; }
            set
            {
                apaxTxtbox[3].Text = value.ToString();
                outData[4] = value;
            }
        }

        public int OutData05
        {
            get { return outData[5]; }
            set
            {
                apaxTxtbox[5].Text = value.ToString();
                outData[5] = value;
            }
        }

        public int OutData06
        {
            get { return outData[6]; }
            set
            {
                apaxTxtbox[6].Text = value.ToString();
                outData[6] = value;
            }
        }

        public int OutData07
        {
            get { return outData[7]; }
            set
            {
                apaxTxtbox[8].Text = value.ToString();
                outData[7] = value;
            }
        }
        public int OutData08
        {
            get { return outData[8]; }
            set
            {
                apaxTxtbox[4].Text = value.ToString();
                outData[8] = value;
            }
        }
        public int OutData09
        {
            get { return outData[9]; }
            set
            {
                apaxTxtbox[7].Text = value.ToString();
                outData[9] = value;
            }
        }
        //20151106 for modbus
        string[] mbData = new string[10];
        public string MBData01
        {
            get { return mbData[0]; }
            set
            {
                modbTxtbox[0].Text = value.ToString();
                mbData[0] = value;
            }
        }
        public string MBData02
        {
            get { return mbData[1]; }
            set
            {
                modbTxtbox[1].Text = value.ToString();
                mbData[1] = value;
            }
        }
        public string MBData03
        {
            get { return mbData[2]; }
            set
            {
                modbTxtbox[2].Text = value.ToString();
                mbData[2] = value;
            }
        }
        public string MBData04
        {
            get { return mbData[3]; }
            set
            {
                modbTxtbox[3].Text = value.ToString();
                mbData[3] = value;
            }
        }
        public string MBData05
        {
            get { return mbData[4]; }
            set
            {
                modbTxtbox[5].Text = value.ToString();
                mbData[4] = value;
            }
        }
        public string MBData06
        {
            get { return mbData[5]; }
            set
            {
                modbTxtbox[6].Text = value.ToString();
                mbData[5] = value;
            }
        }
        public string MBData07
        {
            get { return mbData[6]; }
            set
            {
                modbTxtbox[8].Text = value.ToString();
                mbData[6] = value;
            }
        }
        public string MBData08
        {
            get { return mbData[7]; }
            set
            {
                modbTxtbox[4].Text = value.ToString();
                mbData[7] = value;
            }
        }
        public string MBData09
        {
            get { return mbData[8]; }
            set
            {
                modbTxtbox[7].Text = value.ToString();
                mbData[8] = value;
            }
        }
        //
        bool[] stpchk = new bool[10];
        public bool StpChkIdx0
        {
            get { return stpchk[0]; }
            set
            {
                stpchk[0] = value;
                if (value)
                {
                    StpChkIdx1 = StpChkIdx2 = StpChkIdx3 = StpChkIdx4 = StpChkIdx5
                        = StpChkIdx6 = StpChkIdx7 = StpChkIdx8 = StpChkIdx9 = true;
                }
                else
                {
                    StpChkIdx1 = StpChkIdx2 = StpChkIdx3 = StpChkIdx4 = StpChkIdx5
                        = StpChkIdx6 = StpChkIdx7 = StpChkIdx8 = StpChkIdx9 = false;
                }
            }
        }
        public bool StpChkIdx1
        {
            get { return stpchk[1]; }
            set { stpchk[1] = value;}
        }
        public bool StpChkIdx2
        {
            get { return stpchk[2]; }
            set { stpchk[2] = value; }
        }
        public bool StpChkIdx3
        {
            get { return stpchk[3]; }
            set { stpchk[3] = value;}
        }
        public bool StpChkIdx4
        {
            get { return stpchk[4]; }
            set { stpchk[4] = value; }
        }
        public bool StpChkIdx5
        {
            get { return stpchk[5]; }
            set { stpchk[5] = value; }
        }
        public bool StpChkIdx6
        {
            get { return stpchk[6]; }
            set { stpchk[6] = value;}
        }
        public bool StpChkIdx7
        {
            get { return stpchk[7]; }
            set { stpchk[7] = value;}
        }
        public bool StpChkIdx8
        {
            get { return stpchk[8]; }
            set { stpchk[8] = value; }
        }
        public bool StpChkIdx9
        {
            get { return stpchk[9]; }
            set { stpchk[9] = value;}
        }
        //
        int getdata01 = 0;
        public int GetData01
        {
            get { return getdata01; }
            set { getdata01 = value;
                    getTxtbox[4].Text = getdata01.ToString();/*RaisePropertyChanged("GetData01");*/ }
        }
        int getdata02 = 0;
        public int GetData02
        {
            get { return getdata02; }
            set { getdata02 = value;
                    getTxtbox[7].Text = getdata02.ToString();/*RaisePropertyChanged("GetData02");*/
            }
        }
        #endregion

        public Messenger Messenger
        {
            get
            {
                return Messenger.Instance;
            }
        }

        /// <summary>
        /// main
        /// </summary>
        public WISE_AutoIO_AI_FunTest()
        {
            InitializeComponent();
            //
            //SavefileDialog.DefaultExt = ".txt"; // Default file extension
            //SavefileDialog.Filter = "Text documents (.txt)|*.txt"; // Filter files by extension

            //ModelType = new ObservableCollection<string>
            //{   //For AI used
            //    "Demo",
            //    WISEType.WISE4010LAN.ToString(),
            //    WISEType.WISE4012.ToString(),
            //};
            ChannelType = new ObservableCollection<string>
            {
                "All",
                "0","1","2","3",
            };
            
            VAType = new ObservableCollection<string>
            {   //Note : V_Neg2pt5To2pt5 is not support.
                "V Mode", "A Mode",
                "mV_0To150","mV_0To500","mV_Neg150To150","mV_Neg500To500",
                "V_0To1","V_0To10","V_0To5","V_Neg10To10",
                "V_Neg1To1",/*"V_Neg2pt5To2pt5",*/"V_Neg5To5",
            };

            SelAllChkbox.Checked = StpChkIdx0 = true;//select all steps
            //ProcessView = new ObservableCollection<ListViewData>();
        }

        private void WISE_AutoIO_AI_FunTest_Load(object sender, EventArgs e)
        {
            #region -- Item --
            chkbox.Initialize(); setTxtbox.Initialize();
            getTxtbox.Initialize(); apaxTxtbox.Initialize();
            modbTxtbox.Initialize(); resLabel.Initialize();
            var text_style = new FontFamily("Times New Roman");
            for (int i = 0; i < 9; i++)
            {
                chkbox[i] = new CheckBox();
                chkbox[i].Name = "StpChkIdx" + (i + 1).ToString();
                chkbox[i].Location = new Point(10, 83 + 35 * (i + 1));
                chkbox[i].Text = "";
                chkbox[i].Parent = this;
                chkbox[i].CheckedChanged += new EventHandler(SubChkBoxChanged);

                setTxtbox[i] = new TextBox();
                setTxtbox[i].Size = new Size(60, 25);
                setTxtbox[i].Location = new Point(174, 83 + 35 * (i + 1));                
                setTxtbox[i].Font = new Font(text_style, 12, FontStyle.Regular);
                setTxtbox[i].TextAlign = HorizontalAlignment.Center;
                setTxtbox[i].Parent = this;

                getTxtbox[i] = new TextBox();
                getTxtbox[i].Size = new Size(60, 25);
                getTxtbox[i].Location = new Point(240, 83 + 35 * (i + 1));
                getTxtbox[i].Font = new Font(text_style, 12, FontStyle.Regular);
                getTxtbox[i].TextAlign = HorizontalAlignment.Center;
                getTxtbox[i].Parent = this;

                apaxTxtbox[i] = new TextBox();
                apaxTxtbox[i].Size = new Size(60, 25);
                apaxTxtbox[i].Location = new Point(306, 83 + 35 * (i + 1));
                apaxTxtbox[i].Font = new Font(text_style, 12, FontStyle.Regular);
                apaxTxtbox[i].TextAlign = HorizontalAlignment.Center;
                apaxTxtbox[i].Parent = this;

                modbTxtbox[i] = new TextBox();
                modbTxtbox[i].Size = new Size(60, 25);
                modbTxtbox[i].Location = new Point(372, 83 + 35 * (i + 1));
                modbTxtbox[i].Font = new Font(text_style, 12, FontStyle.Regular);
                modbTxtbox[i].TextAlign = HorizontalAlignment.Center;
                modbTxtbox[i].Parent = this;

                resLabel[i] = new Label();
                resLabel[i].Size = new Size(60, 25);
                resLabel[i].Location = new Point(438, 83 + 35 * (i + 1));
                resLabel[i].Font = new Font(text_style, 12, FontStyle.Regular);
                resLabel[i].Text = "";
                resLabel[i].Parent = this;
            }
            for (int i = 0; i < 9; i++)
            {
                chkbox[i].Checked = true;
            }

            //
            dataGridView1.ColumnHeadersVisible = true;
            DataGridViewTextBoxColumn newCol = new DataGridViewTextBoxColumn(); // add a column to the grid
            newCol.HeaderText = "Time";
            newCol.Name = "clmTs";
            newCol.Visible = true;
            newCol.Width = 20;
            dataGridView1.Columns.Add(newCol);
            //
            newCol = new DataGridViewTextBoxColumn();
            newCol.HeaderText = "Ch";
            newCol.Name = "clmStp";
            newCol.Visible = true;
            newCol.Width = 30;
            dataGridView1.Columns.Add(newCol);
            //
            newCol = new DataGridViewTextBoxColumn();
            newCol.HeaderText = "Step";
            newCol.Name = "clmIns";
            newCol.Visible = true;
            newCol.Width = 100;
            dataGridView1.Columns.Add(newCol);
            //
            newCol = new DataGridViewTextBoxColumn();
            newCol.HeaderText = "Result";
            newCol.Name = "clmDes";
            newCol.Visible = true;
            newCol.Width = 50;
            dataGridView1.Columns.Add(newCol);
            //
            newCol = new DataGridViewTextBoxColumn();
            newCol.HeaderText = "Rowdata";
            newCol.Name = "clmRes";
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
            //20150626 add new service for controller
            HttpReqService = new HttpReqService(Messenger);
            //20151016 ad new service for controller
            ModbusTCPService = new ModbusTCPService();
            //20160316 add for APAX5070
            APAX5070Service = new APAX5070Service();
            //
            GetParaFromFile(); tempIP = ipTxtbox.Text;
            //debug
            //ModcomboBox.SelectedIndex = 1;
            //WISEConnection();
        }
        public void StartTest()//iATester
        {
            if (WISEConnection())
            {
                eStatus(this, new StatusEventArgs(iStatus.Running));
                while(timer.Enabled)
                {
                    Application.DoEvents();
                }

                if (FailCnt > 0)
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
        void ProcessView(ListViewData _obj)
        {
            DataGridViewRow dgvRow;
            DataGridViewCell dgvCell;
            dgvRow = new DataGridViewRow();
            //dgvRow.DefaultCellStyle.Font = new Font(this.Font, FontStyle.Regular);
            if (_obj.Result == "Failed") dgvRow.DefaultCellStyle.ForeColor = Color.Red;
            dgvCell = new DataGridViewTextBoxCell(); //Column Time
            var dataTimeInfo = DateTime.Now.ToString("yyyy-MM-dd HH:MM:ss");
            dgvCell.Value = dataTimeInfo;
            dgvRow.Cells.Add(dgvCell);
            //
            dgvCell = new DataGridViewTextBoxCell();
            dgvCell.Value = _obj.Ch;
            dgvRow.Cells.Add(dgvCell);
            //
            dgvCell = new DataGridViewTextBoxCell();
            dgvCell.Value = _obj.Step;
            dgvRow.Cells.Add(dgvCell);
            //
            dgvCell = new DataGridViewTextBoxCell();
            dgvCell.Value = _obj.Result;
            dgvRow.Cells.Add(dgvCell);
            //
            dgvCell = new DataGridViewTextBoxCell();
            dgvCell.Value = _obj.RowData;
            dgvRow.Cells.Add(dgvCell);

            m_DataGridViewCtrlAddDataRow(dgvRow);
        }
        void GetParaFromFile()
        {
            string sPath = System.Reflection.Assembly.GetAssembly(this.GetType()).Location;
            folderPath = Path.GetDirectoryName(sPath);

            if (File.Exists(folderPath + "\\" + filename))
            {
                using (ExecuteIniClass IniFile = new ExecuteIniClass(Path.Combine(folderPath, filename)))
                {
                    ipTxtbox.Text = IniFile.getKeyValue("Dev", "IP");
                }
            }
        }
        void SetParaToFile()
        {
            if (!File.Exists(folderPath + "\\" + filename))
                File.Create(folderPath + "\\" + filename);

            //save para.
            using (ExecuteIniClass IniFile = new ExecuteIniClass(Path.Combine(folderPath, filename)))
            {
                IniFile.setKeyValue("Dev", "IP", ipTxtbox.Text);
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            switch (mainTaskStp)
            {
                case (int)WISE_AT_AI_Task.wCh_Init:
                    ProcessView(new ListViewData() { Ch = Ref_IO_Mod.Ch, Step = "Initialized", Result = "" });
                    GetDeviceItems(ai_id_offset + Ref_IO_Mod.Ch, out obj);
                    mainTaskStp++;
                    break;
                case (int)WISE_AT_AI_Task.wCh_En:
                    if (StpChkIdx1)
                    {
                        ProcessStep = WISE_AT_AI_Task.wCh_En.ToString();
                        ResStr01 = ChannelEnTest(ai_id_offset + Ref_IO_Mod.Ch) ? "Passed" : "Failed";
                        ProcessView(new ListViewData() { Ch = Ref_IO_Mod.Ch, Step = ProcessStep, Result = ResStr01 });
                    }
                    else mainTaskStp++;
                    break;
                case (int)WISE_AT_AI_Task.wCh_En_Inv:
                    if (StpChkIdx2)
                    {
                        ProcessStep = WISE_AT_AI_Task.wCh_En_Inv.ToString();
                        ChannelEnInvTest(ai_id_offset + Ref_IO_Mod.Ch);
                    }
                    else mainTaskStp++;
                    break;
                case (int)WISE_AT_AI_Task.wCh_En_Inv_res:
                    ProcessView(new ListViewData() { Ch = Ref_IO_Mod.Ch, Step = ProcessStep, Result = ResStr02 });
                    mainTaskStp++;
                    break;
                case (int)WISE_AT_AI_Task.wCh_Wait_En:
                    ProcessStep = WISE_AT_AI_Task.wCh_Wait_En.ToString();
                    EnChProcess(ai_id_offset + Ref_IO_Mod.Ch);
                    mainTaskStp++;
                    break;
                case (int)WISE_AT_AI_Task.wCh_Rng:
                    if (StpChkIdx3)
                    {
                        ProcessStep = WISE_AT_AI_Task.wCh_Rng.ToString();
                        ChannelRangeTest(ai_id_offset + Ref_IO_Mod.Ch);
                    }
                    else mainTaskStp++;
                    break;
                case (int)WISE_AT_AI_Task.wCh_Rng_res:
                    ProcessView(new ListViewData() { Ch = Ref_IO_Mod.Ch, Step = ProcessStep, Result = ResStr03 });
                    mainTaskStp++;
                    break;
                case (int)WISE_AT_AI_Task.wCh_En_Hi_AL_Md0:
                    if (StpChkIdx4)
                    {
                        ProcessStep = WISE_AT_AI_Task.wCh_En_Hi_AL_Md0.ToString();
                        EnHiAMd0Test(ai_id_offset + Ref_IO_Mod.Ch);
                    }
                    else mainTaskStp = (int)WISE_AT_AI_Task.wCh_En_Hi_AL_Md1;
                    break;
                case (int)WISE_AT_AI_Task.wCh_En_Hi_AL_Md0_res:
                    ProcessView(new ListViewData() { Ch = Ref_IO_Mod.Ch, Step = ProcessStep, Result = ResStr04 });
                    mainTaskStp++;
                    break;
                case (int)WISE_AT_AI_Task.wCh_En_Hi_AL_Md1:
                    if (StpChkIdx5)
                    {
                        ProcessStep = WISE_AT_AI_Task.wCh_En_Hi_AL_Md1.ToString();
                        EnHiAMd1Test(ai_id_offset + Ref_IO_Mod.Ch);
                    }
                    else mainTaskStp = (int)WISE_AT_AI_Task.wCh_En_Lo_AL_Md0;
                    break;
                case (int)WISE_AT_AI_Task.wCh_En_Hi_AL_Md1_res:
                    ProcessView(new ListViewData() { Ch = Ref_IO_Mod.Ch, Step = ProcessStep, Result = ResStr05 });
                    mainTaskStp++;
                    break;
                case (int)WISE_AT_AI_Task.wCh_En_Lo_AL_Md0:
                    if (StpChkIdx6)
                    {
                        ProcessStep = WISE_AT_AI_Task.wCh_En_Lo_AL_Md0.ToString();
                        EnLoAMd0Test(ai_id_offset + Ref_IO_Mod.Ch);
                    }
                    else mainTaskStp = (int)WISE_AT_AI_Task.wCh_En_Lo_AL_Md1;
                    break;
                case (int)WISE_AT_AI_Task.wCh_En_Lo_AL_Md0_res:
                    ProcessView(new ListViewData() { Ch = Ref_IO_Mod.Ch, Step = ProcessStep, Result = ResStr06 });
                    mainTaskStp++;
                    break;
                case (int)WISE_AT_AI_Task.wCh_En_Lo_AL_Md1:
                    if (StpChkIdx7)
                    {
                        ProcessStep = WISE_AT_AI_Task.wCh_En_Lo_AL_Md1.ToString();
                        EnLoAMd1Test(ai_id_offset + Ref_IO_Mod.Ch);
                    }
                    else mainTaskStp = (int)WISE_AT_AI_Task.wCh_ChangRng;
                    break;
                case (int)WISE_AT_AI_Task.wCh_En_Lo_AL_Md1_res:
                    ProcessView(new ListViewData() { Ch = Ref_IO_Mod.Ch, Step = ProcessStep, Result = ResStr07 });
                    mainTaskStp++;
                    break;
                case (int)WISE_AT_AI_Task.wCh_ChangRng:
                    ProcessStep = WISE_AT_AI_Task.wCh_ChangRng.ToString();
                    ChangeRng();
                    break;


                case (int)WISE_AT_AI_Task.wFinished:
                    ProcessStep = WISE_AT_AI_Task.wFinished.ToString();
                    ProcessView(new ListViewData() { Ch = Ref_IO_Mod.Ch, Step = "Finished", Result = "" });
                    OutputAO(IntToRowData(RngMod, 0));
                    if (ChSelIdx == 0 && Ref_IO_Mod.Ch < (Ref_IO_Mod.AI_num - 1))
                    {
                        Ref_IO_Mod.Ch++; GetDeviceItems(ai_id_offset + Ref_IO_Mod.Ch, out obj); VerInit();
                    }
                    else
                    {
                        //DelegateHandle.CompProcExe(false);
                        timer.Stop(); RunBtnStr = "Run";
                        if (FailCnt > 0) TestResult = "Fail";
                        else TestResult = "Pass";
                    }
                    break;
                default:

                    mainTaskStp++;
                    break;
            }
        }
        private void SubChkBoxChanged(object sender, EventArgs e)
        {
            var obj = (CheckBox)sender;
            for (int i = 1; i < 10; i++)
            {
                string _name = "StpChkIdx" + i.ToString();
                if (obj.Name == _name)
                    stpchk[i] = obj.Checked;
            }
        }
        private void StartBtn_Click(object sender, EventArgs e)
        {
            if (!timer.Enabled)
            {
                SetParaToFile();
                WISEConnection();
            }
            else
            {
                timer.Stop(); StartBtn.Text = "Run";
            }
        }
        private void idxTxtbox_TextChanged(object sender, EventArgs e)
        {
            _apaxAOslot_idx = Convert.ToInt32(idxTxtbox.Text);
        }
        private void SelAllChkbox_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (SelAllChkbox.Checked)
                {
                    for (int i = 0; i < 9; i++) chkbox[i].Checked = true;
                }
                else
                {
                    for (int i = 0; i < 9; i++) chkbox[i].Checked = false;
                }
            }
            catch
            { }            
        }
        private void ChcomboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            ChSelIdx = ChcomboBox.SelectedIndex;
        }
        private void ModcomboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            VASelIdx = ModcomboBox.SelectedIndex;
        }

        //================================================================================//
        private bool WISEConnection()
        {
            ConnReadyFlg = false;
            Device = new DeviceModel() { IPAddress = ipTxtbox.Text};
            //if (Device.IPAddress == "" || Device.IPAddress == null) Device.IPAddress = tempIP;
            
            if (HttpReqService.HttpReqTCP_Connet(Device))
            {
                Device = HttpReqService.GetDevice();
                ConnReadyFlg = true;
                typTxtbox.Text = Device.ModuleType;
                tssLabel1.Text = "";//error msg
            }
            else
            { tssLabel1.Text = "WISE Disconnected....";}
            //Connect modbus
            if (!ModbusTCPService.ModbusConnection(Device))
                tssLabel1.Text += "Modbus Disconnected....";
            //
            APAX5070 = new DeviceModel() { IPAddress = CDipTxtbox.Text };
            if (APAX5070.IPAddress != "")
            {
                if (!APAX5070Service.ModbusConnection(APAX5070))
                    tssLabel1.Text += "APAX-5070 Disconnected....";
            }
            //
            System.Threading.Thread.Sleep(1000);
            if (!timer.Enabled && ConnReadyFlg)
            {
                if (ProcessInitial())
                {
                    timer.Start(); StartBtn.Text = "Starting";
                }
                else ProcessStep = "Setting Error!";
            }
            else
            {
                timer.Stop(); StartBtn.Text = "Run";
            }
            return ConnReadyFlg;
        }
        private bool ProcessInitial()
        {
            if (Device.ModuleType == "WISE-4012")
            {
                Ref_IO_Mod = new WISE_IO_Model(24012);
                RefTypSelIdx = 3; OutAO_Slot_Idx = 33;
            }
            else if (Device.ModuleType == "WISE-4012E")
            {
                Ref_IO_Mod = new WISE_IO_Model(34012);
                RefTypSelIdx = 2; OutAO_Slot_Idx = 37;
            }
            else// if(Device.ModuleType == "WISE-4010/LAN")
            {
                Ref_IO_Mod = new WISE_IO_Model(4010);
                RefTypSelIdx = 1; OutAO_Slot_Idx = 33;
            }

            _ref_ai_num = Ref_IO_Mod.AI_num;
            lenTxtbox.Text = OutAO_Ch_Len = _ref_ai_num.ToString();
            //
            object _obj;
            GetDeviceItems(ai_id_offset + Ref_IO_Mod.Ch, out _obj);
            var io = (IOListData)_obj;
            if (io.Id == null) return false;
            //
            if (ChSelIdx > 0)
                Ref_IO_Mod.Ch = ChSelIdx - 1;
            else Ref_IO_Mod.Ch = 0;
            //
            FailCnt = 0;
            VerInit();
            return true;
        }
        private void VerInit()
        {
            mainTaskStp = 0;
            ResultIni(); resCnt = 0;
            temp_rng = new int[20]; temp_rng.Initialize();
        }
        private bool GetDeviceItems(int id, out object mod)
        {
            bool flg = false; mod = null;
            while (true)
            {
                try
                {
                    List<IOListData> IO_Data = (List<IOListData>)HttpReqService.GetListOfIOItems("");
                    if (id >= ai_id_offset)
                    {
                        foreach (var item in IO_Data)
                        {
                            if (item.Id == id && item.Ch == id - ai_id_offset)
                            {
                                IOitem = new IOListData()
                                {
                                    Id = item.Id,
                                    Ch = item.Ch,
                                    Tag = item.Tag,
                                    Val = item.Val,
                                    //AI
                                    En = item.En,
                                    Rng = item.Rng,
                                    Evt = item.Evt,
                                    LoA = item.LoA,
                                    HiA = item.HiA,
                                    EgF = item.EgF,
                                    Val_Eg = item.Val_Eg,
                                    cEn = item.cEn,
                                    cRng = item.cRng,
                                    EnLA = item.EnLA,
                                    EnHA = item.EnHA,
                                    LAMd = item.LAMd,
                                    HAMd = item.HAMd,
                                    cLoA = item.cLoA,
                                    cHiA = item.cHiA,
                                    LoS = item.LoS,
                                    HiS = item.HiS,
                                    // add basic
                                    Res = item.Res,
                                    EnB = item.EnB,
                                    BMd = item.BMd,
                                    AiT = item.AiT,
                                    Smp = item.Smp,
                                    AvgM = item.AvgM,
                                    //DI
                                    Md = item.Md,
                                    Inv = item.Inv,
                                    Fltr = item.Fltr,
                                    FtLo = item.FtLo,
                                    FtHi = item.FtHi,
                                    FqT = item.FqT,
                                    FqP = item.FqP,
                                    CntIV = item.CntIV,
                                    CntKp = item.CntKp,
                                    OvLch = item.OvLch,
                                    //DO
                                    FSV = item.FSV,
                                    PsLo = item.PsLo,
                                    PsHi = item.PsHi,
                                    HDT = item.HDT,
                                    LDT = item.LDT,
                                    ACh = item.ACh,
                                    AMd = item.AMd,
                                };
                            }
                        }                        
                    }
                    else if (id >= do_id_offset)
                    {
                        foreach (var item in IO_Data)
                        {
                            if (item.Id == id && item.Ch == id - do_id_offset)
                            {
                                DOitem = new IOListData()
                                {
                                    Id = item.Id,
                                    Ch = item.Ch,
                                    Tag = item.Tag,
                                    Val = item.Val,
                                    //AI
                                    En = item.En,
                                    Rng = item.Rng,
                                    Evt = item.Evt,
                                    LoA = item.LoA,
                                    HiA = item.HiA,
                                    EgF = item.EgF,
                                    Val_Eg = item.Val_Eg,
                                    cEn = item.cEn,
                                    cRng = item.cRng,
                                    EnLA = item.EnLA,
                                    EnHA = item.EnHA,
                                    LAMd = item.LAMd,
                                    HAMd = item.HAMd,
                                    cLoA = item.cLoA,
                                    cHiA = item.cHiA,
                                    LoS = item.LoS,
                                    HiS = item.HiS,
                                    // add basic
                                    Res = item.Res,
                                    EnB = item.EnB,
                                    BMd = item.BMd,
                                    AiT = item.AiT,
                                    Smp = item.Smp,
                                    AvgM = item.AvgM,
                                    //DI
                                    Md = item.Md,
                                    Inv = item.Inv,
                                    Fltr = item.Fltr,
                                    FtLo = item.FtLo,
                                    FtHi = item.FtHi,
                                    FqT = item.FqT,
                                    FqP = item.FqP,
                                    CntIV = item.CntIV,
                                    CntKp = item.CntKp,
                                    OvLch = item.OvLch,
                                    //DO
                                    FSV = item.FSV,
                                    PsLo = item.PsLo,
                                    PsHi = item.PsHi,
                                    HDT = item.HDT,
                                    LDT = item.LDT,
                                    ACh = item.ACh,
                                    AMd = item.AMd,
                                };
                            }
                        }
                    }
                    if (id >= ai_id_offset)
                    {
                        //ViewData = IOitem;
                        mod = IOitem;
                        chiTxtbox.Text = (id - ai_id_offset).ToString();
                        //return true;
                    }
                    else if (id >= do_id_offset) { mod = DOitem;}

                    flg = true; break;
                }
                catch (Exception ex)
                { /*ProcessView.Add(new ListViewData() { RowData = "Open WISE Failed!", });*/
                    string str = ex.ToString();
                }
                
            }
            
            if (flg) return true;

            return false;
        }
        private bool GetMbHA_Flg(int id)//20151112
        {
            try
            {   //Get modbus coil value.
                //var par = CustomController.ReadModbusCoils(DevInfo.DevModbus.HAlm, DevInfo.DevModbus.LenHAlm);
                var par = ModbusTCPService.ReadCoils(Device.MbCoils.HAlm, Device.MbCoils.lenHAlm);
                for (int i = 0; i < par.Length; i++)
                {
                    if (i == id - ai_id_offset)
                    {
                        ProcessView(new ListViewData()
                        {
                            Step = "GetMbHA_Flg",
                            RowData = "Modbus = 0x" + (Device.MbCoils.HAlm + i).ToString("000")
                                + " ; Val: " + par[i].ToString()
                        });
                        return par[i];
                    }
                }
            }
            catch { }
            return false;
        }
        private bool GetMbLA_Flg(int id)//20151112
        {
            try
            {   //Get modbus coil value.
                //var par = CustomController.ReadModbusCoils(DevInfo.DevModbus.LAlm, DevInfo.DevModbus.LenLAlm);
                var par = ModbusTCPService.ReadCoils(Device.MbCoils.LAlm, Device.MbCoils.lenLAlm);
                for (int i = 0; i < par.Length; i++)
                {
                    if (i == id - ai_id_offset)
                    {
                        ProcessView(new ListViewData()
                        {
                            Step = "GetMbLA_Flg",
                            RowData = "Modbus = 0x" + (Device.MbCoils.LAlm + i).ToString("000")
                                + " ; Val: " + par[i].ToString()
                        });
                        return par[i];
                    }
                }
            }
            catch { }
            return false;
        }
        private int GetMbAI_RegVal(int id)//20151112 new
        {
            int res = 0;
            try
            {   //Get modbus reg value.
                //var par = CustomController.ReadModbusRegs(DevInfo.DevModbus.AI, DevInfo.DevModbus.LenAI);
                var par = ModbusTCPService.ReadHoldingRegs(Device.MbRegs.AI, Device.MbRegs.lenAI);
                for (int i = 0; i < par.Length; i++)
                {
                    if (i == id - ai_id_offset)
                    {
                        ProcessView(new ListViewData()
                        {
                            Step = "GetMbAI_RegVal",
                            RowData = "Modbus = 4x" + (Device.MbRegs.AI + i).ToString("000")
                                + " ; Val: " + par[i].ToString()
                        });
                        return par[i];
                    }
                }
            }
            catch { }
            return res;
        }
        private int GetMbAIRng_RegVal(int id)//20151112 new
        {
            int res = 0;
            try
            {   //Get modbus reg value.
                //var par = CustomController.ReadModbusRegs(DevInfo.DevModbus.AICd, DevInfo.DevModbus.LenAICd);
                var par = ModbusTCPService.ReadHoldingRegs(Device.MbRegs.AICd, Device.MbRegs.lenAICd);
                for (int i = 0; i < par.Length; i++)
                {
                    if (i == id - ai_id_offset)
                    {
                        ProcessView(new ListViewData()
                        {
                            Step = "GetMbAIRng_RegVal",
                            RowData = "Modbus = 4x" + (Device.MbRegs.AICd + i).ToString("000")
                                + " ; Val: " + par[i].ToString()
                        });
                        return par[i];
                    }
                }
            }
            catch { }
            return res;
        }

        #region ----  Process Method  -----
        //IOItemViewData IOitem = new IOItemViewData();
        //IOItemViewData PACIOitem = new IOItemViewData();
        IOListData IOitem = new IOListData();
        IOListData DOitem = new IOListData();
        private void ResultIni()
        {
            ResStr01 = ResStr02 = ResStr03 = ResStr04 = ResStr05 = ResStr06 = ResStr07 = ResStr08 = ResStr09 = "";
            SetData01 = SetData02 = SetData03 /*= SetData04 = SetData05 = SetData06 = SetData07*/ = 0;
            SetData04 = SetData05 = SetData06 = SetData07 = "";
            OutData01 = OutData02 = OutData03 = OutData04 = OutData05 = OutData06 = OutData07 = 0;
            MBData01 = MBData02 = MBData03 = MBData04 = MBData05 = MBData06 = MBData07 = "";
            GetData01 = GetData02 = 0;
            ChInvStp = RngChgStp = HiAMd0Stp = HiAMd1Stp = LoAMd0Stp = LoAMd1Stp = 0;
            for (int i = 0; i < 9; i++) getTxtbox[i].Text = "";
        }
        //===================== Function ===============================//
        private bool ChannelEnTest(int id)
        {
            mainTaskStp++;
            GetDeviceItems(id, out obj);
            IOitem.cEn = Ref_IO_Mod.cEn = SetData01 = 1;
            if (IOitem.En != 1)
            {
                //CustomController.UpdateIOConfig(IOitem);
                HttpReqService.UpdateIOConfig(TransferIOModel());
                Thread.Sleep(DelayT);
                GetDeviceItems(id, out obj);
            }
            getTxtbox[0].Text = IOitem.En.ToString();
            ProcessView(new ListViewData() { RowData = "IO_En = " + IOitem.En.ToString() });
            if (IOitem.En > 0 /*&& IOitem.Val > 0*/)
                return true;
            FailCnt++;
            return false;
        }
        int ChInvStp = 0;
        private void ChannelEnInvTest(int id)
        {
            if (ChInvStp >= 99)
            {
                ResStr02 = "Failed"; mainTaskStp++;
                FailCnt++;
            }
            else if (ChInvStp > 1 && ChInvStp < 99)
            {
                ResStr02 = "Passed"; mainTaskStp++;
            }
            else if (ChInvStp == 1)
            {
                GetDeviceItems(id, out obj);
                getTxtbox[1].Text = IOitem.Val.ToString();
                ProcessView(new ListViewData()
                {
                    RowData = "IO_En = " + IOitem.En.ToString()
                        + ", IO_Val = " + IOitem.Val.ToString()
                });
                if (IOitem.En == 0 /*&& IOitem.Val == 0*/)//AI value cannot reset to zero.
                    ChInvStp++;
                else ChInvStp = 99;
            }
            else if (ChInvStp == 0)
            {
                IOitem.cEn = Ref_IO_Mod.cEn = SetData02 = 0;
                HttpReqService.UpdateIOConfig(TransferIOModel());
                //CustomController.UpdateIOConfig(IOitem);
                ChInvStp++;
            }
        }
        //bool reslt = false;
        int resCnt = 0;
        int RngChgStp = 0, Ref_rng = 0;
        private void ChannelRangeTest(int id)
        {
            //20151112 change program to follow step.
            if (RngChgStp >= 99)
            {
                ResStr03 = "Failed"; mainTaskStp++;
                FailCnt++;
            }
            else if (RngChgStp > 5 && RngChgStp < 99)
            {
                ResStr03 = "Passed"; mainTaskStp++;
            }
            else if (RngChgStp == 5)
            {
                //Check modbus AI value would correspond with reading.
                int mb_AIval = GetMbAI_RegVal(id);
                //MBData03 = mb_AIval.ToString();
                GetDeviceItems(id, out obj);//get current value
                //
                ProcessView(new ListViewData()
                {
                    Step = RngChgStp.ToString(),
                    RowData = "mb_AIval = " + mb_AIval.ToString() +
                              ", IOitem.Val = " + IOitem.Val.ToString()
                });
                //
                if (mb_AIval == IOitem.Val
                     || mb_AIval > (IOitem.Val - 100))
                    RngChgStp++;
                else RngChgStp = 99;
            }
            else if (RngChgStp == 4)
            {
                //Check input value would correspond with reading.
                if (DoValDetect(id, Ref_rng))
                    RngChgStp++;
                else RngChgStp = 99;
            }
            else if (RngChgStp == 3)
            {
                //Check modbus rng code would correspond with reading.
                int mb_rng = GetMbAIRng_RegVal(id);
                MBData03 = mb_rng.ToString();

                if (mb_rng == IOitem.Rng) RngChgStp++;
                else RngChgStp = 99;
            }
            else if (RngChgStp == 2)
            {
                //Detect Rng code would correspond with setting.
                if (DoRngDetect(id, Ref_rng))
                    RngChgStp++;
                else RngChgStp = 99;

                temp_rng[resCnt] = Ref_rng;
                resCnt++;
            }
            else if (RngChgStp == 1)
            {
                if (Ref_IO_Mod.AI_num < 0)
                {
                    RngChgStp = 99;
                }
                else
                {
                    foreach (var ref_rng in Ref_IO_Mod.AIRng)
                    {
                        bool flg = false;

                        foreach (var t_rng in temp_rng)
                        {
                            if (ref_rng == t_rng)
                            {
                                flg = true; break;
                            }
                        }

                        if (flg) continue;

                        #region -- Range Selector --
                        if (VASelIdx == 11)
                        {
                            if (ref_rng == (int)ValueRange.mV_0To150
                                || ref_rng == (int)ValueRange.mV_0To500
                                || ref_rng == (int)ValueRange.mV_Neg150To150
                                || ref_rng == (int)ValueRange.mV_Neg500To500
                                || ref_rng == (int)ValueRange.V_0To1
                                || ref_rng == (int)ValueRange.V_0To10
                                || ref_rng == (int)ValueRange.V_0To5
                                || ref_rng == (int)ValueRange.V_Neg10To10
                                || ref_rng == (int)ValueRange.V_Neg1To1
                                || ref_rng == (int)ValueRange.V_Neg2pt5To2pt5
                                //|| ref_rng == (int)ValueRange.V_Neg5To5
                                || ref_rng == (int)ValueRange.mA_4To20
                                || ref_rng == (int)ValueRange.mA_0To20
                                || ref_rng == (int)ValueRange.mA_Neg20To20)
                            {
                                temp_rng[resCnt] = ref_rng;
                                resCnt++; continue;
                            }
                        }
                        //else if (VASelIdx == 11)//V_Neg2pt5To2pt5 not support
                        //{
                        //    if (ref_rng == (int)ValueRange.mV_0To150
                        //        || ref_rng == (int)ValueRange.mV_0To500
                        //        || ref_rng == (int)ValueRange.mV_Neg150To150
                        //        || ref_rng == (int)ValueRange.mV_Neg500To500
                        //        || ref_rng == (int)ValueRange.V_0To1
                        //        || ref_rng == (int)ValueRange.V_0To10
                        //        || ref_rng == (int)ValueRange.V_0To5
                        //        || ref_rng == (int)ValueRange.V_Neg10To10
                        //        || ref_rng == (int)ValueRange.V_Neg1To1
                        //        //|| ref_rng == (int)ValueRange.V_Neg2pt5To2pt5
                        //        || ref_rng == (int)ValueRange.V_Neg5To5
                        //        || ref_rng == (int)ValueRange.mA_4To20
                        //        || ref_rng == (int)ValueRange.mA_0To20
                        //        || ref_rng == (int)ValueRange.mA_Neg20To20)
                        //    {
                        //        temp_rng[resCnt] = ref_rng;
                        //        resCnt++; continue;
                        //    }
                        //}
                        else if (VASelIdx == 10)
                        {
                            if (ref_rng == (int)ValueRange.mV_0To150
                                || ref_rng == (int)ValueRange.mV_0To500
                                || ref_rng == (int)ValueRange.mV_Neg150To150
                                || ref_rng == (int)ValueRange.mV_Neg500To500
                                || ref_rng == (int)ValueRange.V_0To1
                                || ref_rng == (int)ValueRange.V_0To10
                                || ref_rng == (int)ValueRange.V_0To5
                                || ref_rng == (int)ValueRange.V_Neg10To10
                                //|| ref_rng == (int)ValueRange.V_Neg1To1
                                || ref_rng == (int)ValueRange.V_Neg2pt5To2pt5
                                || ref_rng == (int)ValueRange.V_Neg5To5
                                || ref_rng == (int)ValueRange.mA_4To20
                                || ref_rng == (int)ValueRange.mA_0To20
                                || ref_rng == (int)ValueRange.mA_Neg20To20)
                            {
                                temp_rng[resCnt] = ref_rng;
                                resCnt++; continue;
                            }
                        }
                        else if (VASelIdx == 9)
                        {
                            if (ref_rng == (int)ValueRange.mV_0To150
                                || ref_rng == (int)ValueRange.mV_0To500
                                || ref_rng == (int)ValueRange.mV_Neg150To150
                                || ref_rng == (int)ValueRange.mV_Neg500To500
                                || ref_rng == (int)ValueRange.V_0To1
                                || ref_rng == (int)ValueRange.V_0To10
                                || ref_rng == (int)ValueRange.V_0To5
                                //|| ref_rng == (int)ValueRange.V_Neg10To10
                                || ref_rng == (int)ValueRange.V_Neg1To1
                                || ref_rng == (int)ValueRange.V_Neg2pt5To2pt5
                                || ref_rng == (int)ValueRange.V_Neg5To5
                                || ref_rng == (int)ValueRange.mA_4To20
                                || ref_rng == (int)ValueRange.mA_0To20
                                || ref_rng == (int)ValueRange.mA_Neg20To20)
                            {
                                temp_rng[resCnt] = ref_rng;
                                resCnt++; continue;
                            }
                        }
                        else if (VASelIdx == 8)
                        {
                            if (ref_rng == (int)ValueRange.mV_0To150
                                || ref_rng == (int)ValueRange.mV_0To500
                                || ref_rng == (int)ValueRange.mV_Neg150To150
                                || ref_rng == (int)ValueRange.mV_Neg500To500
                                || ref_rng == (int)ValueRange.V_0To1
                                || ref_rng == (int)ValueRange.V_0To10
                                //|| ref_rng == (int)ValueRange.V_0To5
                                || ref_rng == (int)ValueRange.V_Neg10To10
                                || ref_rng == (int)ValueRange.V_Neg1To1
                                || ref_rng == (int)ValueRange.V_Neg2pt5To2pt5
                                || ref_rng == (int)ValueRange.V_Neg5To5
                                || ref_rng == (int)ValueRange.mA_4To20
                                || ref_rng == (int)ValueRange.mA_0To20
                                || ref_rng == (int)ValueRange.mA_Neg20To20)
                            {
                                temp_rng[resCnt] = ref_rng;
                                resCnt++; continue;
                            }
                        }
                        else if (VASelIdx == 7)
                        {
                            if (ref_rng == (int)ValueRange.mV_0To150
                                || ref_rng == (int)ValueRange.mV_0To500
                                || ref_rng == (int)ValueRange.mV_Neg150To150
                                || ref_rng == (int)ValueRange.mV_Neg500To500
                                || ref_rng == (int)ValueRange.V_0To1
                                //|| ref_rng == (int)ValueRange.V_0To10
                                || ref_rng == (int)ValueRange.V_0To5
                                || ref_rng == (int)ValueRange.V_Neg10To10
                                || ref_rng == (int)ValueRange.V_Neg1To1
                                || ref_rng == (int)ValueRange.V_Neg2pt5To2pt5
                                || ref_rng == (int)ValueRange.V_Neg5To5
                                || ref_rng == (int)ValueRange.mA_4To20
                                || ref_rng == (int)ValueRange.mA_0To20
                                || ref_rng == (int)ValueRange.mA_Neg20To20)
                            {
                                temp_rng[resCnt] = ref_rng;
                                resCnt++; continue;
                            }
                        }
                        else if (VASelIdx == 6)
                        {
                            if (ref_rng == (int)ValueRange.mV_0To150
                                || ref_rng == (int)ValueRange.mV_0To500
                                || ref_rng == (int)ValueRange.mV_Neg150To150
                                || ref_rng == (int)ValueRange.mV_Neg500To500
                                //|| ref_rng == (int)ValueRange.V_0To1
                                || ref_rng == (int)ValueRange.V_0To10
                                || ref_rng == (int)ValueRange.V_0To5
                                || ref_rng == (int)ValueRange.V_Neg10To10
                                || ref_rng == (int)ValueRange.V_Neg1To1
                                || ref_rng == (int)ValueRange.V_Neg2pt5To2pt5
                                || ref_rng == (int)ValueRange.V_Neg5To5
                                || ref_rng == (int)ValueRange.mA_4To20
                                || ref_rng == (int)ValueRange.mA_0To20
                                || ref_rng == (int)ValueRange.mA_Neg20To20)
                            {
                                temp_rng[resCnt] = ref_rng;
                                resCnt++; continue;
                            }
                        }
                        else if (VASelIdx == 5)
                        {
                            if (ref_rng == (int)ValueRange.mV_0To150
                                || ref_rng == (int)ValueRange.mV_0To500
                                || ref_rng == (int)ValueRange.mV_Neg150To150
                                //|| ref_rng == (int)ValueRange.mV_Neg500To500
                                || ref_rng == (int)ValueRange.V_0To1
                                || ref_rng == (int)ValueRange.V_0To10
                                || ref_rng == (int)ValueRange.V_0To5
                                || ref_rng == (int)ValueRange.V_Neg10To10
                                || ref_rng == (int)ValueRange.V_Neg1To1
                                || ref_rng == (int)ValueRange.V_Neg2pt5To2pt5
                                || ref_rng == (int)ValueRange.V_Neg5To5
                                || ref_rng == (int)ValueRange.mA_4To20
                                || ref_rng == (int)ValueRange.mA_0To20
                                || ref_rng == (int)ValueRange.mA_Neg20To20)
                            {
                                temp_rng[resCnt] = ref_rng;
                                resCnt++; continue;
                            }
                        }
                        else if (VASelIdx == 4)
                        {
                            if (ref_rng == (int)ValueRange.mV_0To150
                                || ref_rng == (int)ValueRange.mV_0To500
                                //|| ref_rng == (int)ValueRange.mV_Neg150To150
                                || ref_rng == (int)ValueRange.mV_Neg500To500
                                || ref_rng == (int)ValueRange.V_0To1
                                || ref_rng == (int)ValueRange.V_0To10
                                || ref_rng == (int)ValueRange.V_0To5
                                || ref_rng == (int)ValueRange.V_Neg10To10
                                || ref_rng == (int)ValueRange.V_Neg1To1
                                || ref_rng == (int)ValueRange.V_Neg2pt5To2pt5
                                || ref_rng == (int)ValueRange.V_Neg5To5
                                || ref_rng == (int)ValueRange.mA_4To20
                                || ref_rng == (int)ValueRange.mA_0To20
                                || ref_rng == (int)ValueRange.mA_Neg20To20)
                            {
                                temp_rng[resCnt] = ref_rng;
                                resCnt++; continue;
                            }
                        }
                        else if (VASelIdx == 3)
                        {
                            if (ref_rng == (int)ValueRange.mV_0To150
                                //|| ref_rng == (int)ValueRange.mV_0To500
                                || ref_rng == (int)ValueRange.mV_Neg150To150
                                || ref_rng == (int)ValueRange.mV_Neg500To500
                                || ref_rng == (int)ValueRange.V_0To1
                                || ref_rng == (int)ValueRange.V_0To10
                                || ref_rng == (int)ValueRange.V_0To5
                                || ref_rng == (int)ValueRange.V_Neg10To10
                                || ref_rng == (int)ValueRange.V_Neg1To1
                                || ref_rng == (int)ValueRange.V_Neg2pt5To2pt5
                                || ref_rng == (int)ValueRange.V_Neg5To5
                                || ref_rng == (int)ValueRange.mA_4To20
                                || ref_rng == (int)ValueRange.mA_0To20
                                || ref_rng == (int)ValueRange.mA_Neg20To20)
                            {
                                temp_rng[resCnt] = ref_rng;
                                resCnt++; continue;
                            }
                        }
                        else if (VASelIdx == 2)
                        {
                            if (//ref_rng == (int)ValueRange.mV_0To150
                                ref_rng == (int)ValueRange.mV_0To500
                                || ref_rng == (int)ValueRange.mV_Neg150To150
                                || ref_rng == (int)ValueRange.mV_Neg500To500
                                || ref_rng == (int)ValueRange.V_0To1
                                || ref_rng == (int)ValueRange.V_0To10
                                || ref_rng == (int)ValueRange.V_0To5
                                || ref_rng == (int)ValueRange.V_Neg10To10
                                || ref_rng == (int)ValueRange.V_Neg1To1
                                || ref_rng == (int)ValueRange.V_Neg2pt5To2pt5
                                || ref_rng == (int)ValueRange.V_Neg5To5
                                || ref_rng == (int)ValueRange.mA_4To20
                                || ref_rng == (int)ValueRange.mA_0To20
                                || ref_rng == (int)ValueRange.mA_Neg20To20)
                            {
                                temp_rng[resCnt] = ref_rng;
                                resCnt++; continue;
                            }
                        }
                        else if (VASelIdx == 1)
                        {
                            VModNum = 10;
                            if (ref_rng == (int)ValueRange.mV_0To150
                                || ref_rng == (int)ValueRange.mV_0To500
                                || ref_rng == (int)ValueRange.mV_Neg150To150
                                || ref_rng == (int)ValueRange.mV_Neg500To500
                                || ref_rng == (int)ValueRange.V_0To1
                                || ref_rng == (int)ValueRange.V_0To10
                                || ref_rng == (int)ValueRange.V_0To5
                                || ref_rng == (int)ValueRange.V_Neg10To10
                                || ref_rng == (int)ValueRange.V_Neg1To1
                                || ref_rng == (int)ValueRange.V_Neg2pt5To2pt5
                                || ref_rng == (int)ValueRange.V_Neg5To5)
                            {
                                temp_rng[resCnt] = ref_rng;
                                resCnt++; continue;
                            }
                        }
                        else
                        {
                            AModNum = 3;
                            if (ref_rng == (int)ValueRange.mA_4To20
                                || ref_rng == (int)ValueRange.mA_0To20
                                || ref_rng == (int)ValueRange.mA_Neg20To20)
                            {
                                temp_rng[resCnt] = ref_rng;
                                resCnt++; continue;
                            }
                        }
                        #endregion

                        //還沒切換過的Rng
                        if (DoChangeRange(ref_rng))
                        {
                            RngMod = ref_rng;
                            //需要控制PAC端AO輸出
                            OutData03 = OutputAO(IntToRowData(ref_rng, Ref_IO_Mod.APAX_AO_HiVal));
                            //Thread.Sleep(DelayT);
                            Ref_rng = ref_rng;
                        }
                        break;
                    }
                }
                RngChgStp++;
            }
            else if (RngChgStp == 0)
            {
                ResStr03 = ResStr04 = ResStr05 = ResStr06 = ResStr07 = ResStr08 = ResStr09
                    = MBData03 = MBData04 = MBData05 = MBData06 = MBData07 = MBData08 = MBData09 = "";
                HiAMd0Stp = HiAMd1Stp = LoAMd0Stp = LoAMd1Stp = 0;
                Ref_rng = 0;
                RngChgStp++;
            }
        }

        int HiAMd0Stp = 0;
        private bool EnHiAMd0Test(int id)
        {
            if (HiAMd0Stp >= 99)
            {
                ResStr04 = "Failed"; mainTaskStp++;
                FailCnt++;
            }
            else if (HiAMd0Stp > 3 && HiAMd0Stp < 99)
            {
                ResStr04 = "Passed"; mainTaskStp++;
            }
            else if (HiAMd0Stp == 3)
            {
                GetDeviceItems(id, out obj);
                getTxtbox[3].Text = IOitem.HiA.ToString();
                ProcessView(new ListViewData() { Step = "3", RowData = "HiA = " + IOitem.HiA.ToString() 
                    + ", Val_Eg = " + IOitem.EgF.ToString() });
                //Check modbus HA would correspond with reading.20151112
                bool mb_ha = GetMbHA_Flg(id);
                MBData04 = mb_ha.ToString();

                if (IOitem.HiA < 1 && !mb_ha)
                {
                    HiAMd0Stp++;
                }
                else HiAMd0Stp = 99;
            }
            else if (HiAMd0Stp == 2)
            {
                GetDeviceItems(id, out obj);
                ProcessView(new ListViewData() { Step = "2", RowData = "HiA = " + IOitem.HiA.ToString() 
                    + ", Val_Eg = " + IOitem.EgF.ToString() });
                //Check modbus HA would correspond with reading.20151112
                bool mb_ha = GetMbHA_Flg(id);
                MBData04 = mb_ha.ToString();

                if (IOitem.HiA > 0 && mb_ha)
                {
                    if (GetDOConfHA())//20150803 add do alarm function.
                    {
                        HiAMd0Stp++;
                    }
                    else HiAMd0Stp = 99;
                    //需要控制PAC端AO輸出
                    OutData04 = OutputAO(IntToRowData(RngMod, Ref_IO_Mod.APAX_AO_MdVal));
                }
                else HiAMd0Stp = 99;
            }
            else if (HiAMd0Stp == 1)
            {
                //需要控制PAC端AO輸出
                OutData04 = OutputAO(IntToRowData(RngMod, Ref_IO_Mod.APAX_AO_HiVal));
                HiAMd0Stp++;
            }
            else if (HiAMd0Stp == 0)
            {
                IOitem.EnHA = Ref_IO_Mod.EnHA = 1;//enable
                IOitem.EnLA = Ref_IO_Mod.EnLA = 0;//disable
                IOitem.HAMd = Ref_IO_Mod.HAMd = 0;//mode 0
                SetData04 = IOitem.cHiA.ToString();//IOitem.cHiA
                HttpReqService.UpdateIOConfig(TransferIOModel());
                //CustomController.UpdateIOConfig(IOitem);
                SetDOConfHA();//20150803 add do alarm function.
                //
                IOitem.HiA = 0;
                HttpReqService.UpdateIOValue(TransferIO_Val_Model());
                //CustomController.UpdateIOValue(IOitem);
                HiAMd0Stp++;
            }

            return false;
        }

        int HiAMd1Stp = 0;
        private bool EnHiAMd1Test(int id)
        {
            if (HiAMd1Stp >= 99)
            {
                ResStr05 = "Failed"; mainTaskStp++;
                FailCnt++;
            }
            else if (HiAMd1Stp > 3 && HiAMd1Stp < 99)
            {
                ResStr05 = "Passed"; mainTaskStp++;
            }
            else if (HiAMd1Stp == 3)
            {
                GetDeviceItems(id, out obj);
                getTxtbox[5].Text = IOitem.HiA.ToString();
                ProcessView(new ListViewData() { Step = "3", RowData = "IO_HiA = " + IOitem.HiA.ToString() 
                    + ", Val_Eg = " + IOitem.EgF.ToString() });
                //Check modbus HA would correspond with reading.20151112
                bool mb_ha = GetMbHA_Flg(id);
                MBData05 = mb_ha.ToString();

                if (IOitem.HiA > 0 && mb_ha)
                {
                    HiAMd1Stp++;
                }
                else HiAMd1Stp = 99;
            }
            else if (HiAMd1Stp == 2)
            {
                GetDeviceItems(id, out obj);
                ProcessView(new ListViewData() { Step = "2", RowData = "IO_HiA = " + IOitem.HiA.ToString() 
                    + ", Val_Eg = " + IOitem.EgF.ToString() });
                //Check modbus HA would correspond with reading.20151112
                bool mb_ha = GetMbHA_Flg(id);
                MBData05 = mb_ha.ToString();

                if (IOitem.HiA > 0 && mb_ha)
                {
                    //需要控制PAC端AO輸出
                    OutData05 = OutputAO(IntToRowData(RngMod, Ref_IO_Mod.APAX_AO_MdVal));
                    HiAMd1Stp++;
                }
                else HiAMd1Stp = 99;
            }
            else if (HiAMd1Stp == 1)
            {
                //需要控制PAC端AO輸出
                OutData05 = OutputAO(IntToRowData(RngMod, Ref_IO_Mod.APAX_AO_HiVal));
                HiAMd1Stp++;
            }
            else if (HiAMd1Stp == 0)
            {
                IOitem.EnHA = Ref_IO_Mod.EnHA = 1;//enable
                IOitem.EnLA = Ref_IO_Mod.EnLA = 0;//disable
                IOitem.HAMd = Ref_IO_Mod.HAMd = 1;//mode 0
                SetData05 = IOitem.cHiA.ToString();
                HttpReqService.UpdateIOConfig(TransferIOModel());
                //CustomController.UpdateIOConfig(IOitem);
                //
                IOitem.HiA = 0;
                HttpReqService.UpdateIOValue(TransferIO_Val_Model());
                //CustomController.UpdateIOValue(IOitem);
                //
                //Thread.Sleep(DelayT);
                //GetDeviceItems(id);
                HiAMd1Stp++;
            }

            return false;
        }

        int LoAMd0Stp = 0;
        private bool EnLoAMd0Test(int id)
        {
            if (LoAMd0Stp >= 99)
            {
                ResStr06 = "Failed"; mainTaskStp++;
                FailCnt++;
            }
            else if (LoAMd0Stp > 3 && LoAMd0Stp < 99)
            {
                ResStr06 = "Passed"; mainTaskStp++;
            }
            else if (LoAMd0Stp == 3)
            {
                GetDeviceItems(id, out obj);
                getTxtbox[6].Text = IOitem.LoA.ToString();
                ProcessView(new ListViewData() { Step = "3", RowData = "LoA = " + IOitem.LoA.ToString() 
                    + ", Val_Eg = " + IOitem.EgF.ToString() });
                //Check modbus HA would correspond with reading.20151112
                bool mb_la = GetMbLA_Flg(id);
                MBData06 = mb_la.ToString();

                if (IOitem.LoA < 1 && !mb_la)
                {
                    LoAMd0Stp++;
                }
                else LoAMd0Stp = 99;
            }
            else if (LoAMd0Stp == 2)
            {
                GetDeviceItems(id, out obj);
                ProcessView(new ListViewData() { Step = "2", RowData = "IO_LoA = " + IOitem.LoA.ToString() 
                    + ", Val_Eg = " + IOitem.EgF.ToString() });
                //Check modbus HA would correspond with reading.20151112
                bool mb_la = GetMbLA_Flg(id);
                MBData06 = mb_la.ToString();

                if (IOitem.LoA > 0 && mb_la)
                {
                    if (GetDOConfLA())//20150803 add do alarm function.
                    {
                        LoAMd0Stp++;
                    }
                    else LoAMd0Stp = 99;
                    //需要控制PAC端AO輸出
                    OutData06 = OutputAO(IntToRowData(RngMod, Ref_IO_Mod.APAX_AO_MdVal));
                    //Thread.Sleep(DelayT);
                }
                else LoAMd0Stp = 99;
            }
            else if (LoAMd0Stp == 1)
            {
                //需要控制PAC端AO輸出
                OutData06 = OutputAO(IntToRowData(RngMod, Ref_IO_Mod.APAX_AO_LoVal));
                LoAMd0Stp++;
            }
            else if (LoAMd0Stp == 0)
            {
                IOitem.EnHA = Ref_IO_Mod.EnHA = 0;//disable
                IOitem.EnLA = Ref_IO_Mod.EnLA = 1;//enable
                IOitem.LAMd = Ref_IO_Mod.LAMd = 0;//mode 0
                SetData06 = IOitem.cLoA.ToString();
                //CustomController.UpdateIOConfig(IOitem);
                HttpReqService.UpdateIOConfig(TransferIOModel());
                SetDOConfLA();//20150803 add do alarm function.
                //
                IOitem.LoA = 0; IOitem.HiA = 0;
                HttpReqService.UpdateIOValue(TransferIO_Val_Model());
                //CustomController.UpdateIOValue(IOitem);
                LoAMd0Stp++;
            }

            return false;
        }

        int LoAMd1Stp = 0;
        private bool EnLoAMd1Test(int id)
        {
            if (LoAMd1Stp >= 99)
            {
                ResStr07 = "Failed"; mainTaskStp++;
                FailCnt++;
            }
            else if (LoAMd1Stp > 3 && LoAMd1Stp < 99)
            {
                ResStr07 = "Passed"; mainTaskStp++;
            }
            else if (LoAMd1Stp == 3)
            {
                GetDeviceItems(id, out obj);
                getTxtbox[8].Text = IOitem.LoA.ToString();
                ProcessView(new ListViewData() { Step = "3", RowData = "IO_LoA = " + IOitem.LoA.ToString() + ", Val_Eg = " + IOitem.EgF.ToString() });
                //Check modbus HA would correspond with reading.20151112
                bool mb_la = GetMbLA_Flg(id);
                MBData07 = mb_la.ToString();

                if (IOitem.LoA > 0 && mb_la)
                {
                    LoAMd1Stp++;
                }
                else LoAMd1Stp = 99;
            }
            else if (LoAMd1Stp == 2)
            {
                GetDeviceItems(id, out obj);
                ProcessView(new ListViewData() { Step = "2", RowData = "IO_LoA = " + IOitem.LoA.ToString() + ", Val_Eg = " + IOitem.EgF.ToString() });
                //Check modbus HA would correspond with reading.20151112
                bool mb_la = GetMbLA_Flg(id);
                MBData07 = mb_la.ToString();

                if (IOitem.LoA > 0 && mb_la)
                {
                    //需要控制PAC端AO輸出
                    OutData07 = OutputAO(IntToRowData(RngMod, Ref_IO_Mod.APAX_AO_MdVal));
                    LoAMd1Stp++;
                }
                else LoAMd1Stp = 99;
            }
            else if (LoAMd1Stp == 1)
            {
                //需要控制PAC端AO輸出
                OutData07 = OutputAO(IntToRowData(RngMod, Ref_IO_Mod.APAX_AO_LoVal));
                LoAMd1Stp++;
            }
            else if (LoAMd1Stp == 0)
            {
                IOitem.LAMd = Ref_IO_Mod.LAMd = 1;//mode 1
                SetData07 = IOitem.cLoA.ToString();
                HttpReqService.UpdateIOConfig(TransferIOModel());
                //CustomController.UpdateIOConfig(IOitem);
                //
                IOitem.LoA = 0;
                HttpReqService.UpdateIOValue(TransferIO_Val_Model());
                //CustomController.UpdateIOValue(IOitem);
                LoAMd1Stp++;
            }

            return false;
        }

        private void ChangeRng()
        {
            if (resCnt < Ref_IO_Mod.AIRng.Length)
            {
                if (StpChkIdx8 || StpChkIdx9) ResetDO();

                if (StpChkIdx3)
                {
                    //20151112 needs to check V/A mode number.
                    if (VASelIdx == 0 && resCnt < (Ref_IO_Mod.AIRng.Length - AModNum)
                         || VASelIdx == 1 && resCnt < Ref_IO_Mod.AIRng.Length)
                    {
                        mainTaskStp = (int)WISE_AT_AI_Task.wCh_Rng;
                    }
                    else
                        mainTaskStp = (int)WISE_AT_AI_Task.wFinished;
                }
                else mainTaskStp = (int)WISE_AT_AI_Task.wFinished;

                RngChgStp = 0;//20151112 reset step
            }
            else
            {
                mainTaskStp++;
            }
        }

        private void EnChProcess(int id)
        {
            GetDeviceItems(id, out obj);
            IOitem.cEn = Ref_IO_Mod.cEn = 1;
            HttpReqService.UpdateIOConfig(TransferIOModel());
            //CustomController.UpdateIOConfig(IOitem);
            //Thread.Sleep(DelayT);
            //GetDeviceItems(id);
        }








        //private void ResetDefault()
        //{
        //    IOitem.Md = 0; IOitem.Inv = 0;
        //    HttpReqService.UpdateIOConfig(TransferIOModel());
        //    //CustomController.UpdateIOConfig(IOitem);
        //    OutputDO(false);
        //}
        private void DoATRunComplete()
        {
            bool res = FailCnt > 0 ? false : true;
            //DelegateHandle.CompProcExe(res);
        }
        private IOModel TransferIOModel()
        {
            IOModel item = new IOModel()
            {
                //Basic
                Id = IOitem.Id,
                Ch = IOitem.Ch,
                Tag = IOitem.Tag,
                Val = IOitem.Val,
                //AI
                Val_Eg = IOitem.Val_Eg,
                cEn = IOitem.cEn,
                cRng = IOitem.cRng,
                EnLA = IOitem.EnLA,
                EnHA = IOitem.EnHA,
                LAMd = IOitem.LAMd,
                HAMd = IOitem.HAMd,
                cLoA = IOitem.cLoA,
                cHiA = IOitem.cHiA,
                LoS = IOitem.LoS,
                HiS = IOitem.HiS,
                // add basic
                Res = IOitem.Res,
                EnB = IOitem.EnB,
                BMd = IOitem.BMd,
                AiT = IOitem.AiT,
                Smp = IOitem.Smp,
                AvgM = IOitem.AvgM,
                //DI
                Md = IOitem.Md,
                Inv = IOitem.Inv,
                Fltr = IOitem.Fltr,
                FtLo = IOitem.FtLo,
                FtHi = IOitem.FtHi,
                FqT = IOitem.FqT,
                FqP = IOitem.FqP,
                CntIV = IOitem.CntIV,
                CntKp = IOitem.CntKp,
                //DO
                FSV = IOitem.FSV,
                PsLo = IOitem.PsLo,
                PsHi = IOitem.PsHi,
                HDT = IOitem.HDT,
                LDT = IOitem.LDT,
                ACh = IOitem.ACh,
                AMd = IOitem.AMd,
            };
            return item;
        }
        private IOModel TransferIO_Val_Model()
        {
            IOModel item = new IOModel()
            {
                //Basic
                Id = IOitem.Id,
                Ch = IOitem.Ch,
                //AI
                LoA = IOitem.LoA,
                HiA = IOitem.HiA,
                //DO
                Val = IOitem.Val,
                PsCtn = IOitem.PsCtn,
                PsStop = IOitem.PsStop,
                PsIV = IOitem.PsIV,
                //DI
                Cnting = IOitem.Cnting,
                ClrCnt = IOitem.ClrCnt,
                OvLch = IOitem.OvLch,
            };
            return item;
        }
        //For DO
        private IOModel TransferIOModelByListData(IOListData _obj)
        {
            IOModel item = new IOModel()
            {
                //Basic
                Id = _obj.Id,
                Ch = _obj.Ch,
                Tag = _obj.Tag,
                Val = _obj.Val,
                //AI
                Val_Eg = _obj.Val_Eg,
                cEn = _obj.cEn,
                cRng = _obj.cRng,
                EnLA = _obj.EnLA,
                EnHA = _obj.EnHA,
                LAMd = _obj.LAMd,
                HAMd = _obj.HAMd,
                cLoA = _obj.cLoA,
                cHiA = _obj.cHiA,
                LoS = _obj.LoS,
                HiS = _obj.HiS,
                // add basic
                Res = _obj.Res,
                EnB = _obj.EnB,
                BMd = _obj.BMd,
                AiT = _obj.AiT,
                Smp = _obj.Smp,
                AvgM = _obj.AvgM,
                //DI
                Md = _obj.Md,
                Inv = _obj.Inv,
                Fltr = _obj.Fltr,
                FtLo = _obj.FtLo,
                FtHi = _obj.FtHi,
                FqT = _obj.FqT,
                FqP = _obj.FqP,
                CntIV = _obj.CntIV,
                CntKp = _obj.CntKp,
                //DO
                FSV = _obj.FSV,
                PsLo = _obj.PsLo,
                PsHi = _obj.PsHi,
                HDT = _obj.HDT,
                LDT = _obj.LDT,
                ACh = _obj.ACh,
                AMd = _obj.AMd,
            };
            return item;
        }
        private IOModel TransferIO_Val_ModelByListData(IOListData _obj)
        {
            IOModel item = new IOModel()
            {
                //Basic
                Id = _obj.Id,
                Ch = _obj.Ch,
                //AI
                LoA = _obj.LoA,
                HiA = _obj.HiA,
                //DO
                Val = _obj.Val,
                PsCtn = _obj.PsCtn,
                PsStop = _obj.PsStop,
                PsIV = _obj.PsIV,
                //DI
                Cnting = _obj.Cnting,
                ClrCnt = _obj.ClrCnt,
                OvLch = _obj.OvLch,
            };
            return item;
        }
        #endregion

        //----------------------------------------------------------//
        void UpdateCh_Typ(ObservableCollection<string> obj)
        {
            ChcomboBox.Items.Clear();
            foreach (var _str in obj)
            {
                ChcomboBox.Items.Add(_str);
            }
            ChcomboBox.SelectedIndex = 0;
        }
        void UpdateVATyp(ObservableCollection<string> obj)
        {
            ModcomboBox.Items.Clear();
            foreach (var _typ in obj)
            {
                ModcomboBox.Items.Add(_typ);
            }
            ModcomboBox.SelectedIndex = 0;
        }

        //private int OutputDO(bool TF)
        //{
        //    //if (CtrlDevSelIdx == 0)
        //        OutputAPAX5070(TF);
        //    //else
        //        //OutputPAC(TF);

        //    return TF ? 1 : 0;
        //}
        //private void OutputAPAX5070(bool TF)
        //{
        //    int val = TF ? (int)1 : 0;
        //    int idx = OutDO_Slot_Idx + Ref_IO_Mod.Ch;
        //    APAX5070Service.ForceSigCoil(idx, val);
        //    Thread.Sleep(100);
        //}
        private bool DoChangeRange(int _rng)
        {
            SetData03 = Ref_IO_Mod.cRng = _rng;
            //Set Initial HA/LA Value and AO output value.
            #region -- Range --
            if (_rng == (int)ValueRange.mA_4To20 || _rng == (int)ValueRange.mA_0To20)
            {
                IOitem.cHiA = Ref_IO_Mod.cHiA = "15.000";
                IOitem.cLoA = Ref_IO_Mod.cLoA = "6.000";
                //
                Ref_IO_Mod.APAX_AO_HiVal = 20000;
                Ref_IO_Mod.APAX_AO_MdVal = 10000;
                Ref_IO_Mod.APAX_AO_LoVal = 4000;
                ProcessView(new ListViewData()
                {
                    RowData = "Change IO_Rng = mA_4To20 || mA_0To20"
                });
            }
            else if (_rng == (int)ValueRange.mA_Neg20To20)
            {
                //Note : APAX-5028 only support 0~20mA.
                IOitem.cHiA = Ref_IO_Mod.cHiA = "15.000";
                IOitem.cLoA = Ref_IO_Mod.cLoA = "2.000";
                //
                Ref_IO_Mod.APAX_AO_HiVal = 20000;
                Ref_IO_Mod.APAX_AO_MdVal = 10000;
                Ref_IO_Mod.APAX_AO_LoVal = 0;
                ProcessView(new ListViewData()
                {
                    RowData = "Change IO_Rng = "
                        + ValueRange.mA_Neg20To20.ToString()
                });
            }
            else if (_rng == (int)ValueRange.mV_0To150)
            {
                IOitem.cHiA = Ref_IO_Mod.cHiA = "130.000";
                IOitem.cLoA = Ref_IO_Mod.cLoA = "50.000";
                //
                Ref_IO_Mod.APAX_AO_HiVal = 150;
                Ref_IO_Mod.APAX_AO_MdVal = 100;
                Ref_IO_Mod.APAX_AO_LoVal = 0;
                ProcessView(new ListViewData()
                {
                    RowData = "Change IO_Rng = "
                        + ValueRange.mV_0To150.ToString()
                });
            }
            else if (_rng == (int)ValueRange.mV_0To500)
            {
                IOitem.cHiA = Ref_IO_Mod.cHiA = "400.000";
                IOitem.cLoA = Ref_IO_Mod.cLoA = "100.000";
                //
                Ref_IO_Mod.APAX_AO_HiVal = 500;
                Ref_IO_Mod.APAX_AO_MdVal = 300;
                Ref_IO_Mod.APAX_AO_LoVal = 0;
                ProcessView(new ListViewData()
                {
                    RowData = "Change IO_Rng = "
                        + ValueRange.mV_0To500.ToString()
                });
            }
            else if (_rng == (int)ValueRange.V_0To1)
            {
                IOitem.cHiA = Ref_IO_Mod.cHiA = "0.700";
                IOitem.cLoA = Ref_IO_Mod.cLoA = "0.300";
                //
                Ref_IO_Mod.APAX_AO_HiVal = 1000;
                Ref_IO_Mod.APAX_AO_MdVal = 500;
                Ref_IO_Mod.APAX_AO_LoVal = 0;
                ProcessView(new ListViewData()
                {
                    RowData = "Change IO_Rng = "
                        + ValueRange.V_0To1.ToString()
                });
            }
            else if (_rng == (int)ValueRange.V_0To5)
            {
                IOitem.cHiA = Ref_IO_Mod.cHiA = "4.000";
                IOitem.cLoA = Ref_IO_Mod.cLoA = "2.000";
                //
                Ref_IO_Mod.APAX_AO_HiVal = 5000;
                Ref_IO_Mod.APAX_AO_MdVal = 3000;
                Ref_IO_Mod.APAX_AO_LoVal = 0;
                ProcessView(new ListViewData()
                {
                    RowData = "Change IO_Rng = "
                        + ValueRange.V_0To5.ToString()
                });
            }
            else if (_rng == (int)ValueRange.V_0To10)
            {
                IOitem.cHiA = Ref_IO_Mod.cHiA = "7.000";
                IOitem.cLoA = Ref_IO_Mod.cLoA = "3.000";
                //
                Ref_IO_Mod.APAX_AO_HiVal = 10000;
                Ref_IO_Mod.APAX_AO_MdVal = 5000;
                Ref_IO_Mod.APAX_AO_LoVal = 1000;
                ProcessView(new ListViewData()
                {
                    RowData = "Change IO_Rng = "
                        + ValueRange.V_0To10.ToString()
                });
            }
            else if (_rng == (int)ValueRange.mV_Neg150To150)
            {
                IOitem.cHiA = Ref_IO_Mod.cHiA = "130.000";
                IOitem.cLoA = Ref_IO_Mod.cLoA = "-130.000";
                //
                Ref_IO_Mod.APAX_AO_HiVal = 150;
                Ref_IO_Mod.APAX_AO_MdVal = 100;
                Ref_IO_Mod.APAX_AO_LoVal = -150;
                ProcessView(new ListViewData()
                {
                    RowData = "Change IO_Rng = "
                        + ValueRange.mV_Neg150To150.ToString()
                });
            }
            else if (_rng == (int)ValueRange.mV_Neg500To500)
            {
                IOitem.cHiA = Ref_IO_Mod.cHiA = "400.000";
                IOitem.cLoA = Ref_IO_Mod.cLoA = "-400.000";
                //
                Ref_IO_Mod.APAX_AO_HiVal = 500;
                Ref_IO_Mod.APAX_AO_MdVal = 100;
                Ref_IO_Mod.APAX_AO_LoVal = -500;
                ProcessView(new ListViewData()
                {
                    RowData = "Change IO_Rng = "
                        + ValueRange.mV_Neg500To500.ToString()
                });
            }
            else if (_rng == (int)ValueRange.V_Neg1To1)
            {
                IOitem.cHiA = Ref_IO_Mod.cHiA = "0.700";
                IOitem.cLoA = Ref_IO_Mod.cLoA = "-0.700";
                //
                Ref_IO_Mod.APAX_AO_HiVal = 1000;
                Ref_IO_Mod.APAX_AO_MdVal = 500;
                Ref_IO_Mod.APAX_AO_LoVal = -1000;
                ProcessView(new ListViewData()
                {
                    RowData = "Change IO_Rng = "
                        + ValueRange.V_Neg1To1.ToString()
                });
            }
            else if (_rng == (int)ValueRange.V_Neg2pt5To2pt5)
            {
                ProcessView(new ListViewData()
                {
                    RowData = "Change IO_Rng = "
                        + ValueRange.V_Neg2pt5To2pt5.ToString()
                });
            }
            else if (_rng == (int)ValueRange.V_Neg5To5)
            {
                IOitem.cHiA = Ref_IO_Mod.cHiA = "3.000";
                IOitem.cLoA = Ref_IO_Mod.cLoA = "-3.000";
                //
                Ref_IO_Mod.APAX_AO_HiVal = 5000;
                Ref_IO_Mod.APAX_AO_MdVal = 1000;
                Ref_IO_Mod.APAX_AO_LoVal = -5000;
                ProcessView(new ListViewData()
                {
                    RowData = "Change IO_Rng = "
                        + ValueRange.V_Neg5To5.ToString()
                });
            }
            else if (_rng == (int)ValueRange.V_Neg10To10)
            {
                IOitem.cHiA = Ref_IO_Mod.cHiA = "9.000";
                IOitem.cLoA = Ref_IO_Mod.cLoA = "-9.000";
                //
                Ref_IO_Mod.APAX_AO_HiVal = 10000;
                Ref_IO_Mod.APAX_AO_MdVal = 1;
                Ref_IO_Mod.APAX_AO_LoVal = -10000;
                ProcessView(new ListViewData()
                {
                    RowData = "Change IO_Rng = "
                        + ValueRange.V_Neg10To10.ToString()
                });
            }
            #endregion

            //if (IOitem.Rng != _rng)
            {
                IOitem.cRng = Ref_IO_Mod.cRng = _rng;
                HttpReqService.UpdateIOConfig(TransferIOModel());
                //CustomController.UpdateIOConfig(IOitem);
            }
            return true;
        }

        private bool DoRngDetect(int _id, int _rng)
        {
            GetDeviceItems(_id, out obj);
            ProcessView(new ListViewData()
            {
                Step = RngChgStp.ToString(),
                RowData = "IO_Rng = " + IOitem.Rng.ToString() +
                          ", cRng = " + Ref_IO_Mod.cRng.ToString()
            });
            getTxtbox[2].Text = IOitem.Rng.ToString();
            if (IOitem.Rng == Ref_IO_Mod.cRng) return true;


            return false;
        }

        private bool DoValDetect(int _id, int _rng)
        {
            GetDeviceItems(_id, out obj);
            //
            #region -- Range --
            ProcessView(new ListViewData()
            {
                Step = RngChgStp.ToString(),
                RowData = "EgF = " + IOitem.EgF.ToString() +
                          ", AO_HiVal = " + Ref_IO_Mod.APAX_AO_HiVal.ToString()
            });
            //用AI輸入的最大值作為判斷
            //WISE-4012 vA101B00 之後改為EgF
            if (_rng == (int)ValueRange.mA_4To20 || _rng == (int)ValueRange.mA_0To20
                 || _rng == (int)ValueRange.mA_Neg20To20)
            {
                if (IOitem.Rng == Ref_IO_Mod.cRng &&
                    IOitem.EgF > Ref_IO_Mod.APAX_AO_HiVal / 1000 - 5) return true;
            }
            else if (_rng == (int)ValueRange.mV_0To150 || _rng == (int)ValueRange.mV_Neg150To150)
            {
                if (IOitem.Rng == Ref_IO_Mod.cRng &&
                    IOitem.EgF > Ref_IO_Mod.APAX_AO_HiVal - 10) return true;
            }
            else if (_rng == (int)ValueRange.mV_0To500 || _rng == (int)ValueRange.mV_Neg500To500)
            {
                if (IOitem.Rng == Ref_IO_Mod.cRng &&
                    IOitem.EgF > Ref_IO_Mod.APAX_AO_HiVal - 100) return true;
            }
            else if (_rng == (int)ValueRange.V_0To10
                     || _rng == (int)ValueRange.V_0To5 || _rng == (int)ValueRange.V_Neg10To10
                     || _rng == (int)ValueRange.V_Neg5To5
                     || _rng == (int)ValueRange.V_Neg2pt5To2pt5)
            {
                if (IOitem.Rng == Ref_IO_Mod.cRng &&
                    IOitem.EgF > Ref_IO_Mod.APAX_AO_HiVal - 1000) return true;
            }
            else if (_rng == (int)ValueRange.V_0To1 || _rng == (int)ValueRange.V_Neg1To1)
            {
                if (IOitem.Rng == Ref_IO_Mod.cRng &&
                    IOitem.EgF > Ref_IO_Mod.APAX_AO_HiVal - 100) return true;
            }
            //20161101 fix wise4010 problem
            if (Device.ModuleType == "WISE-4010/LAN")//DevInfo.MainDev.ModuleType == "WISE-4010/LAN"
            {
                if (IOitem.Rng == Ref_IO_Mod.cRng &&
                    IOitem.Eg > Ref_IO_Mod.APAX_AO_HiVal / 1000 - 5) return true;
            }

            #endregion

            return false;
        }

        private void ResetDO()
        {
            object _obj;
            GetDeviceItems(do_id_offset + DOChSet, out _obj);
            DOitem = (IOListData)_obj;
            Ref_IO_Mod.Md = DOitem.Md = 0;
            Ref_IO_Mod.ACh = DOitem.ACh = DOChSet;
            Ref_IO_Mod.AMd = DOitem.AMd = 0;
            HttpReqService.UpdateIOConfig(TransferIOModelByListData(DOitem));
            //CustomController.UpdateIOConfig(DOitem);
            Thread.Sleep(500);
            DOitem.Val = 0;
            HttpReqService.UpdateIOValue(TransferIO_Val_ModelByListData(DOitem));
            //CustomController.UpdateIOValue(DOitem);
        }
        int DOChSet = 0;
        private void SetDOConfHA()
        {
            if (!StpChkIdx8) return;

            //需要注意DO與AI的數量比較                
            if (Ref_IO_Mod.Ch < Ref_IO_Mod.DO_num) DOChSet = Ref_IO_Mod.Ch;
            else DOChSet = Ref_IO_Mod.Ch - Ref_IO_Mod.DO_num;
            object _obj;
            GetDeviceItems(do_id_offset + DOChSet, out _obj);
            DOitem = (IOListData)_obj;
            Ref_IO_Mod.Md = DOitem.Md = 4;
            Ref_IO_Mod.ACh = DOitem.ACh = Ref_IO_Mod.Ch;
            Ref_IO_Mod.AMd = DOitem.AMd = SetData08 = 1;
            HttpReqService.UpdateIOConfig(TransferIOModelByListData(DOitem));
            //CustomController.UpdateIOConfig(DOitem);
            Thread.Sleep(100);
            DOitem.Val = 0;
            HttpReqService.UpdateIOValue(TransferIO_Val_ModelByListData(DOitem));
            //CustomController.UpdateIOValue(DOitem);

            ProcessView(new ListViewData()
            {
                Ch = DOitem.Ch,
                Step = "SetDOConfHA",
                RowData = "DO.Val = " + DOitem.Val.ToString()
                    + " DOChSet = " + DOChSet.ToString()
                    + " Ref_IO_Mod.Ch = " + Ref_IO_Mod.Ch.ToString()
            });
        }

        private bool GetDOConfHA()
        {
            if (StpChkIdx8)
            {
                object _obj;
                GetDeviceItems(do_id_offset + DOChSet, out _obj);
                DOitem = (IOListData)_obj;
                //getTxtbox[4].Text
                GetData01 = DOitem.AMd.Value;
                ProcessView(new ListViewData()
                {
                    Ch = DOitem.Ch,
                    Step = "GetDOConfHA",
                    Result = ResStr08,
                    RowData = "DO.Val = " + DOitem.Val.ToString()
                });
                if (DOitem.Val < 1)
                {
                    ResStr08 = "Failed";
                    return false;
                }
                ResStr08 = "Passed";
            }

            return true;
        }

        private void SetDOConfLA()
        {
            if (StpChkIdx9)
            {
                //需要注意DO與AI的數量比較                
                if (Ref_IO_Mod.Ch < Ref_IO_Mod.DO_num) DOChSet = Ref_IO_Mod.Ch;
                else DOChSet = Ref_IO_Mod.Ch - Ref_IO_Mod.DO_num;

                object _obj;
                GetDeviceItems(do_id_offset + DOChSet, out _obj);
                DOitem = (IOListData)_obj;
                Ref_IO_Mod.Md = DOitem.Md = 4;
                Ref_IO_Mod.ACh = DOitem.ACh = Ref_IO_Mod.Ch;
                Ref_IO_Mod.AMd = DOitem.AMd = SetData09 = 2;
                HttpReqService.UpdateIOConfig(TransferIOModelByListData(DOitem));
                //CustomController.UpdateIOConfig(DOitem);
                Thread.Sleep(100);
                DOitem.Val = 0;
                HttpReqService.UpdateIOValue(TransferIO_Val_ModelByListData(DOitem));
                //CustomController.UpdateIOValue(DOitem);

                ProcessView(new ListViewData()
                {
                    Ch = DOitem.Ch,
                    Step = "SetDOConfLA",
                    RowData = "DO.Val = " + DOitem.Val.ToString()
                        + " DOChSet = " + DOChSet.ToString()
                        + " Ref_IO_Mod.Ch = " + Ref_IO_Mod.Ch.ToString()
                });
            }
        }

        private bool GetDOConfLA()
        {
            if (StpChkIdx9)
            {
                object _obj;
                GetDeviceItems(do_id_offset + DOChSet, out _obj);
                DOitem = (IOListData)_obj;
                //getTxtbox[7].Text 
                GetData02 = DOitem.AMd.Value;
                ProcessView(new ListViewData()
                {
                    Ch = DOitem.Ch,
                    Step = "GetDOConfLA",
                    Result = ResStr09,
                    RowData = "DO.Val = " + DOitem.Val.ToString()
                });
                if (DOitem.Val < 1)
                {
                    ResStr09 = "Failed";
                    return false;
                }
                ResStr09 = "Passed";
            }

            return true;
        }

        void RefreshChTypeItem()
        {
            //APAX-5028 Slot1
            //if (CtrlDevSelIdx == 1)
            //    OutAO_Slot_Idx = 37;
            //else
            //    OutAO_Slot_Idx = 33;
        }

        //DeviceInfo CtrlDevInfo;
        private bool InitRefCtrlDev()
        {
            //if (APAX_Demo) return true;
            //try
            //{
            //    if (CtrlDevSelIdx == 0)
            //    {
            //        CtrlDevInfo = CustomController.GetDeviceItemViewData(DevBase.APAX5070, true);//透過service更新資料
            //        if (CtrlDevInfo.MainDev.ModuleType != "")
            //        {
            //            return true;
            //        }
            //    }
            //    else if (!APAX_Demo && CtrlDevSelIdx == 1)
            //    {
            //        CtrlDevInfo = CustomController.GetDeviceItemViewData(DevBase.APAX_Ctrllor, true);//透過service更新資料
            //        if (CtrlDevInfo.MainDev.ModuleType != "")
            //        {
            //            return true;
            //        }
            //    }
            //}
            //catch { ProcessView(new ListViewData() { RowData = "Open Controller Device Failed!", }); }

            return false;
        }

        private uint IntToRowData(int mod, int val)//20151111 new
        {
            uint res = 0;
            float var = (float)val / 1000;
            if (mod == (int)ValueRange.mA_4To20
                                || mod == (int)ValueRange.mA_0To20
                                || mod == (int)ValueRange.mA_Neg20To20)
            {
                res = (uint)(var * 65535 / 20);
            }
            else
                res = (uint)((var + 10) * 65535 / 20);
            return res;
        }

        private int RowDataToInt(uint val)//20151111 new
        {
            int res = 0;
            float var = (float)val;
            if (IOitem.cRng == (int)ValueRange.V_0To10
                    || IOitem.cRng == (int)ValueRange.V_0To5
                    || IOitem.cRng == (int)ValueRange.V_Neg10To10
                    || IOitem.cRng == (int)ValueRange.V_Neg5To5
                    || IOitem.cRng == (int)ValueRange.V_Neg2pt5To2pt5)
            {
                res = (int)(var / (float)65535 * (float)20 - 10);
            }
            else
                res = (int)((var / (float)65535 * (float)20 - 10) * 1000);
            return res;
        }

        private int OutputAO(uint row)
        {
            //if (CtrlDevSelIdx == 0)
                OutputAPAX5070(row);
            //else if (APAX_Demo)
            //    return 0;
            //else
            //    OutputPAC(row);

            return (int)row;//TF ? 1 : 0;
        }

        private int OutputPAC(uint row)
        {
            //PACIOitem.Id = OutAO_Slot_Idx;
            //PACIOitem.Ch = Ref_IO_Mod.Ch;
            //PACIOitem.Val = row;
            //if (APAX_Demo) return RowDataToInt(row);
            //CustomController.UpdateIOConfig(DevBase.APAX_Ctrllor, PACIOitem);

            //return RowDataToInt(row);
            return 0;
        }

        private void OutputAPAX5070(uint row)
        {
            //int val = TF ? (int)1 : 0;
            int idx = OutAO_Slot_Idx + Ref_IO_Mod.Ch;
            APAX5070Service.ForceSigReg(idx, (int)row);
            //CustomController.OutputModbusSingleRegs(idx, (int)row);
            Thread.Sleep(100);
        }



    }//class
    public class ListViewData
    {
        public int Ch { get; set; }
        public string Step { get; set; }
        public string Result { get; set; }
        public string RowData { get; set; }
    }
    public enum WISE_AT_AI_Task
    {
        wCh_Init = 1,
        wCh_En = 2, //channel enable function test
        wCh_En_Inv = 3, //channel enable inverse function test
        wCh_En_Inv_res = 4,
        wCh_Wait_En = 5,
        wCh_Rng = 6, //channel range function test
        wCh_Rng_res = 7,
        wCh_En_Hi_AL_Md0 = 8,//20150803 combine DO ALDrived function test.
        wCh_En_Hi_AL_Md0_res = 9,
        wCh_En_Hi_AL_Md1 = 10,
        wCh_En_Hi_AL_Md1_res = 11,
        wCh_En_Lo_AL_Md0 = 12,//20150803 combine DO ALDrived function test.
        wCh_En_Lo_AL_Md0_res = 13,
        wCh_En_Lo_AL_Md1 = 14,
        wCh_En_Lo_AL_Md1_res = 15,
        wCh_ChangRng = 16,



        wFinished = 20,


    }
    public class ExecuteIniClass : IDisposable
    {
        [DllImport("kernel32")]
        private static extern long WritePrivateProfileString(string section, string key, string val, string filePath);
        [DllImport("kernel32")]
        private static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retVal, int size, string filePath);

        private bool bDisposed = false;
        private string _FilePath = string.Empty;
        public string FilePath
        {
            get
            {
                if (_FilePath == null)
                    return string.Empty;
                else
                    return _FilePath;
            }
            set
            {
                if (_FilePath != value)
                    _FilePath = value;
            }
        }

        /// <summary>
        /// 建構子。
        /// </summary>
        /// <param name="path">檔案路徑。</param>      
        public ExecuteIniClass(string path)
        {
            _FilePath = path;
        }

        /// <summary>
        /// 解構子。
        /// </summary>
        ~ExecuteIniClass()
        {
            Dispose(false);
        }

        /// <summary>
        /// 釋放資源(程式設計師呼叫)。
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this); //要求系統不要呼叫指定物件的完成項。
        }

        /// <summary>
        /// 釋放資源(給系統呼叫的)。
        /// </summary>        
        protected virtual void Dispose(bool IsDisposing)
        {
            if (bDisposed)
            {
                return;
            }
            if (IsDisposing)
            {
                //補充：

                //這裡釋放具有實做 IDisposable 的物件(資源關閉或是 Dispose 等..)
                //ex: DataSet DS = new DataSet();
                //可在這邊 使用 DS.Dispose();
                //或是 DS = null;
                //或是釋放 自訂的物件。
                //因為我沒有這類的物件，故意留下這段 code ;若繼承這個類別，
                //可覆寫這個函式。
            }

            bDisposed = true;
        }


        /// <summary>
        /// 設定 KeyValue 值。
        /// </summary>
        /// <param name="IN_Section">Section。</param>
        /// <param name="IN_Key">Key。</param>
        /// <param name="IN_Value">Value。</param>
        public void setKeyValue(string IN_Section, string IN_Key, string IN_Value)
        {
            WritePrivateProfileString(IN_Section, IN_Key, IN_Value, this._FilePath);
        }

        /// <summary>
        /// 取得 Key 相對的 Value 值。
        /// </summary>
        /// <param name="IN_Section">Section。</param>
        /// <param name="IN_Key">Key。</param>        
        public string getKeyValue(string IN_Section, string IN_Key)
        {
            StringBuilder temp = new StringBuilder(255);
            int i = GetPrivateProfileString(IN_Section, IN_Key, "", temp, 255, this._FilePath);
            return temp.ToString();
        }



        /// <summary>
        /// 取得 Key 相對的 Value 值，若沒有則使用預設值(DefaultValue)。
        /// </summary>
        /// <param name="Section">Section。</param>
        /// <param name="Key">Key。</param>
        /// <param name="DefaultValue">DefaultValue。</param>        
        public string getKeyValue(string Section, string Key, string DefaultValue)
        {
            StringBuilder sbResult = null;
            try
            {
                sbResult = new StringBuilder(255);
                GetPrivateProfileString(Section, Key, "", sbResult, 255, this._FilePath);
                return (sbResult.Length > 0) ? sbResult.ToString() : DefaultValue;
            }
            catch
            {
                return string.Empty;
            }
        }
    }
}
