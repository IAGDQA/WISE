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

namespace WISE_AutoIO_DI_FunTest
{
    public partial class WISE_AutoIO_DI_FunTest : Form, iATester.iCom
    {
        CheckBox[] chkbox = new CheckBox[7];
        TextBox[] setTxtbox = new TextBox[7];
        TextBox[] getTxtbox = new TextBox[7];
        TextBox[] apaxTxtbox = new TextBox[7];
        TextBox[] modbTxtbox = new TextBox[7];
        Label[] resLabel = new Label[7];

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
        int di_id_offset = 0;
        int DelayT = 500;
        bool APAX_Demo, AutoRunFlg;
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
        int _apaxDOslot_idx = 1;//for modbus/apax5580
        public int OutDO_Slot_Idx
        {
            get { return _apaxDOslot_idx; }
            set
            {
                _apaxDOslot_idx = value;
                idxTxtbox.Text = _apaxDOslot_idx.ToString();
                //RaisePropertyChanged("OutDO_Slot_Idx");
            }

        }
        int _ref_di_num = 0;
        public string OutDO_Ch_Len
        {
            get
            { return "0 ~ " + (_ref_di_num - 1).ToString(); }
            set
            {
                //RaisePropertyChanged("OutDO_Ch_Len");
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
        public int SetData04
        {
            get { return setData[4]; }
            set
            {
                setTxtbox[3].Text = value.ToString();
                setData[4] = value;
            }
        }
        public int SetData05
        {
            get { return setData[5]; }
            set
            {
                setTxtbox[4].Text = value.ToString();
                setData[5] = value;
            }
        }
        public int SetData06
        {
            get { return setData[6]; }
            set
            {
                setTxtbox[5].Text = value.ToString();
                setData[6] = value;
            }
        }
        public int SetData07
        {
            get { return setData[7]; }
            set
            {
                setTxtbox[6].Text = value.ToString();
                setData[7] = value;
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
                resLabel[4].Text = value.ToString();
                resStr[5] = value;
            }
        }
        public string ResStr06
        {
            get { return resStr[6]; }
            set
            {
                resLabel[5].Text = value.ToString();
                resStr[6] = value;
            }
        }
        public string ResStr07
        {
            get { return resStr[7]; }
            set
            {
                resLabel[6].Text = value.ToString();
                resStr[7] = value;
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
                apaxTxtbox[4].Text = value.ToString();
                outData[5] = value;
            }
        }

        public int OutData06
        {
            get { return outData[6]; }
            set
            {
                apaxTxtbox[5].Text = value.ToString();
                outData[6] = value;
            }
        }

        public int OutData07
        {
            get { return outData[7]; }
            set
            {
                apaxTxtbox[6].Text = value.ToString();
                outData[7] = value;
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
                modbTxtbox[4].Text = value.ToString();
                mbData[4] = value;
            }
        }
        public string MBData06
        {
            get { return mbData[5]; }
            set
            {
                modbTxtbox[5].Text = value.ToString();
                mbData[5] = value;
            }
        }
        public string MBData07
        {
            get { return mbData[6]; }
            set
            {
                modbTxtbox[6].Text = value.ToString();
                mbData[6] = value;
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
        public WISE_AutoIO_DI_FunTest()
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
            { "Digital"};
            //VAType = new ObservableCollection<string>
            //{   //Note : V_Neg2pt5To2pt5 is not support.
            //    "V Mode", "A Mode",
            //    "mV_0To150","mV_0To500","mV_Neg150To150","mV_Neg500To500",
            //    "V_0To1","V_0To10","V_0To5","V_Neg10To10",
            //    "V_Neg1To1",/*"V_Neg2pt5To2pt5",*/"V_Neg5To5",
            //};

            SelAllChkbox.Checked = StpChkIdx0 = true;//select all steps
            //ProcessView = new ObservableCollection<ListViewData>();
        }

        private void WISE_AutoIO_DI_FunTest_Load(object sender, EventArgs e)
        {
            #region -- Item --
            chkbox.Initialize(); setTxtbox.Initialize();
            getTxtbox.Initialize(); apaxTxtbox.Initialize();
            modbTxtbox.Initialize(); resLabel.Initialize();
            var text_style = new FontFamily("Times New Roman");
            for (int i = 0; i < 7; i++)
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
            for (int i = 0; i < 7; i++)
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
                    label23.Text = "iA running....";
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
            Application.DoEvents(); label23.Text = "iA finished....";
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
                case (int)WISE_AT_DI_Task.wCh_Init:
                    //ProcessView.Add(new ListViewData() { Ch = Ref_IO_Mod.Ch, Step = "Initialized", Result = "" });
                    ProcessView(new ListViewData() { Ch = Ref_IO_Mod.Ch, Step = "Initialized", Result = "" });
                    GetDeviceItems(di_id_offset + Ref_IO_Mod.Ch);
                    ResetDefault();
                    mainTaskStp++;
                    break;
                case (int)WISE_AT_DI_Task.wCh_DI_Fun:
                    if (StpChkIdx1)
                    {
                        GetDeviceItems(di_id_offset + Ref_IO_Mod.Ch);
                        ProcessStep = WISE_AT_DI_Task.wCh_DI_Fun.ToString();
                        ChDI_input(di_id_offset + Ref_IO_Mod.Ch);
                    }
                    else mainTaskStp++;
                    break;
                case (int)WISE_AT_DI_Task.wCh_DI_Fun_res:
                    //ProcessView.Add(new ListViewData() { Ch = Ref_IO_Mod.Ch, Step = ProcessStep, Result = ResStr01 });
                    ProcessView(new ListViewData() { Ch = Ref_IO_Mod.Ch, Step = ProcessStep, Result = ResStr01 });
                    mainTaskStp++;
                    break;
                case (int)WISE_AT_DI_Task.wCh_DI_Inv:
                    if (StpChkIdx2)
                    {
                        ProcessStep = WISE_AT_DI_Task.wCh_DI_Inv.ToString();
                        ChDI_INV_Fun(di_id_offset + Ref_IO_Mod.Ch);
                    }
                    else mainTaskStp = (int)WISE_AT_DI_Task.wCh_DI_RstToInit;
                    break;
                case (int)WISE_AT_DI_Task.wCh_DI_Inv_res:
                    //ProcessView.Add(new ListViewData() { Ch = Ref_IO_Mod.Ch, Step = ProcessStep, Result = ResStr02 });
                    ProcessView(new ListViewData() { Ch = Ref_IO_Mod.Ch, Step = ProcessStep, Result = ResStr02 });
                    mainTaskStp++;
                    break;
                case (int)WISE_AT_DI_Task.wCh_DI_RstToInit:
                    ProcessStep = WISE_AT_DI_Task.wCh_DI_RstToInit.ToString();
                    ResetDefault(); mainTaskStp++;
                    break;
                case (int)WISE_AT_DI_Task.wCh_DI_LtoH_Latch:
                    if (StpChkIdx3)
                    {
                        ProcessStep = WISE_AT_DI_Task.wCh_DI_LtoH_Latch.ToString();
                        ChDI_LtoH_latch(di_id_offset + Ref_IO_Mod.Ch);
                    }
                    else mainTaskStp = (int)WISE_AT_DI_Task.wCh_DI_HtoL_Latch;
                    break;
                case (int)WISE_AT_DI_Task.wCh_DI_LtoH_Latch_res:
                    //ProcessView.Add(new ListViewData() { Ch = Ref_IO_Mod.Ch, Step = ProcessStep, Result = ResStr03 });
                    ProcessView(new ListViewData() { Ch = Ref_IO_Mod.Ch, Step = ProcessStep, Result = ResStr03 });
                    mainTaskStp++;
                    break;
                case (int)WISE_AT_DI_Task.wCh_DI_HtoL_Latch:
                    if (StpChkIdx4)
                    {
                        ProcessStep = WISE_AT_DI_Task.wCh_DI_HtoL_Latch.ToString();
                        ChDI_HtoL_latch(di_id_offset + Ref_IO_Mod.Ch);
                    }
                    else mainTaskStp = (int)WISE_AT_DI_Task.wCh_DI_Counter;
                    break;
                case (int)WISE_AT_DI_Task.wCh_DI_HtoL_Latch_res:
                    //ProcessView.Add(new ListViewData() { Ch = Ref_IO_Mod.Ch, Step = ProcessStep, Result = ResStr04 });
                    ProcessView(new ListViewData() { Ch = Ref_IO_Mod.Ch, Step = ProcessStep, Result = ResStr04 });
                    mainTaskStp++;
                    break;
                case (int)WISE_AT_DI_Task.wCh_DI_Counter:
                    if (StpChkIdx5)
                    {
                        ProcessStep = WISE_AT_DI_Task.wCh_DI_Counter.ToString();
                        ChDI_Ctr_Fun(di_id_offset + Ref_IO_Mod.Ch);
                    }
                    else mainTaskStp = (int)WISE_AT_DI_Task.wCh_DI_Counter_Stup;
                    break;
                case (int)WISE_AT_DI_Task.wCh_DI_Counter_res:
                    //ProcessView.Add(new ListViewData() { Ch = Ref_IO_Mod.Ch, Step = ProcessStep, Result = ResStr05 });
                    ProcessView(new ListViewData() { Ch = Ref_IO_Mod.Ch, Step = ProcessStep, Result = ResStr05 });
                    mainTaskStp++;
                    break;
                case (int)WISE_AT_DI_Task.wCh_DI_Counter_Stup:
                    if (StpChkIdx6)
                    {
                        ProcessStep = WISE_AT_DI_Task.wCh_DI_Counter_Stup.ToString();
                        ChDI_CtrST_Fun(di_id_offset + Ref_IO_Mod.Ch);
                    }
                    else mainTaskStp = (int)WISE_AT_DI_Task.wFinished;
                    break;
                case (int)WISE_AT_DI_Task.wCh_DI_Counter_Stup_res:
                    //ProcessView.Add(new ListViewData() { Ch = Ref_IO_Mod.Ch, Step = ProcessStep, Result = ResStr06 });
                    ProcessView(new ListViewData() { Ch = Ref_IO_Mod.Ch, Step = ProcessStep, Result = ResStr06 });
                    mainTaskStp++;
                    break;


                case (int)WISE_AT_DI_Task.wFinished:
                    ProcessStep = WISE_AT_DI_Task.wFinished.ToString();
                    //ProcessView.Add(new ListViewData() { Ch = Ref_IO_Mod.Ch, Step = "Finished", Result = "" });
                    ProcessView(new ListViewData() { Ch = Ref_IO_Mod.Ch, Step = "Finished", Result = "" });
                    if (ChSelIdx == 0 && Ref_IO_Mod.Ch < (Ref_IO_Mod.DI_num - 1))
                    {
                        ProcessView(new ListViewData() { Ch = Ref_IO_Mod.Ch, Step = "Change next channel", Result = "" });
                        Ref_IO_Mod.Ch++; GetDeviceItems(di_id_offset + Ref_IO_Mod.Ch); VerInit();
                    }
                    else
                    {
                        timer.Stop(); RunBtnStr = "Run";
                        //
                        if (AutoRunFlg) DoATRunComplete();
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
                SetParaToFile(); label22.Text = "0";
                WISEConnection();
            }
            else
            {
                timer.Stop(); StartBtn.Text = "Run";
            }
        }
        private void idxTxtbox_TextChanged(object sender, EventArgs e)
        {
            OutDO_Slot_Idx = Convert.ToInt32(idxTxtbox.Text);
        }
        private void SelAllChkbox_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (SelAllChkbox.Checked)
                {
                    for (int i = 0; i < 7; i++) chkbox[i].Checked = true;
                }
                else
                {
                    for (int i = 0; i < 7; i++) chkbox[i].Checked = false;
                }
            }
            catch
            { }            
        }

        private void ChcomboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            ChSelIdx = ChcomboBox.SelectedIndex;
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
            label22.Text = "1";
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
            if (Device.ModuleType == "WISE-4051")
            {
                Ref_IO_Mod = new WISE_IO_Model(24051);
                //RefTypSelIdx = 6;
                OutDO_Slot_Idx = 9;
            }
            else if (Device.ModuleType == "WISE-4012")
            {
                Ref_IO_Mod = new WISE_IO_Model(24012);
                //RefTypSelIdx = 5;
                OutDO_Slot_Idx = 1;//
            }
            else if (Device.ModuleType == "WISE-4012E")
            {
                Ref_IO_Mod = new WISE_IO_Model(34012);
                //RefTypSelIdx = 5;
                OutDO_Slot_Idx = 17;
            }
            else if (Device.ModuleType == "WISE-4060")
            {
                Ref_IO_Mod = new WISE_IO_Model(4060);
                //RefTypSelIdx = 4;
                OutDO_Slot_Idx = 5;
            }
            else if (Device.ModuleType == "WISE-4050")
            {
                Ref_IO_Mod = new WISE_IO_Model(4050);
                //RefTypSelIdx = 3;
                OutDO_Slot_Idx = 1;
            }
            else if (Device.ModuleType == "WISE-4060/LAN")
            {
                Ref_IO_Mod = new WISE_IO_Model(4060);
                //RefTypSelIdx = 2;
                OutDO_Slot_Idx = 133;
            }
            else if (Device.ModuleType == "WISE-4050/LAN")
            {
                Ref_IO_Mod = new WISE_IO_Model(4050);
                //RefTypSelIdx = 1;
                OutDO_Slot_Idx = 129;
            }
            else
            {
                APAX_Demo = true;
                Ref_IO_Mod = new WISE_IO_Model(24051);
            }

            _ref_di_num = Ref_IO_Mod.DI_num;
            lenTxtbox.Text = OutDO_Ch_Len = _ref_di_num.ToString();
            //
            if (!GetDeviceItems(di_id_offset + Ref_IO_Mod.Ch)
                || !InitRefCtrlDev()) return false;
            //
            if (ChSelIdx > 0)
                Ref_IO_Mod.Ch = ChSelIdx - 1;
            else Ref_IO_Mod.Ch = 0;
            //
            FailCnt = 0;
            VerInit(); label22.Text = "2";
            return true;
        }
        private void VerInit()
        {
            TestResult = "N/A";
            mainTaskStp = 0;
            ResultIni();
        }
        private bool GetDeviceItems(int id)
        {
            bool flg = false;int WDT = 0; label22.Text = "3";
            while (true)
            {
                if(WDT > 3000)
                {
                    tssLabel1.Text += "Get Data Error.......";break;
                }
                WDT++;
                label22.Text = "4"; Application.DoEvents();
                try
                {
                    List<IOListData> IO_Data = (List<IOListData>)HttpReqService.GetListOfIOItems("");
                    foreach (var item in IO_Data)
                    {
                        if (item.Id == id && item.Ch == id - di_id_offset)
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
                    //foreach (var item in IO_Data)
                    //{
                    //    if (item.Id == id && item.Ch == id - di_id_offset)
                    //    {
                    //        IOitem = new IOItemViewData()
                    //        {
                    //            IO_Id = item.Id,
                    //            IO_Ch = item.Ch,
                    //            IO_Tag = item.Tag,
                    //            IO_Val = item.Val,
                    //            //AI
                    //            IO_En = item.En,
                    //            IO_Rng = item.Rng,
                    //            IO_Evt = item.Evt,
                    //            IO_LoA = item.LoA,
                    //            IO_HiA = item.HiA,
                    //            IO_EgF = item.EgF,
                    //            IO_Val_Eg = item.Val_Eg,
                    //            IO_cEn = item.cEn,
                    //            IO_cRng = item.cRng,
                    //            IO_EnLA = item.EnLA,
                    //            IO_EnHA = item.EnHA,
                    //            IO_LAMd = item.LAMd,
                    //            IO_HAMd = item.HAMd,
                    //            IO_cLoA = item.cLoA,
                    //            IO_cHiA = item.cHiA,
                    //            IO_LoS = item.LoS,
                    //            IO_HiS = item.HiS,
                    //            // add basic
                    //            IO_Res = item.Res,
                    //            IO_EnB = item.EnB,
                    //            IO_BMd = item.BMd,
                    //            IO_AiT = item.AiT,
                    //            IO_Smp = item.Smp,
                    //            IO_AvgM = item.AvgM,
                    //            //DI
                    //            IO_Md = item.Md,
                    //            IO_Inv = item.Inv,
                    //            IO_Fltr = item.Fltr,
                    //            IO_FtLo = item.FtLo,
                    //            IO_FtHi = item.FtHi,
                    //            IO_FqT = item.FqT,
                    //            IO_FqP = item.FqP,
                    //            IO_CntIV = item.CntIV,
                    //            IO_CntKp = item.CntKp,
                    //            IO_OvLch = item.OvLch,
                    //            //DO
                    //            IO_FSV = item.FSV,
                    //            IO_PsLo = item.PsLo,
                    //            IO_PsHi = item.PsHi,
                    //            IO_HDT = item.HDT,
                    //            IO_LDT = item.LDT,
                    //            IO_ACh = item.ACh,
                    //            IO_AMd = item.AMd,
                    //        };
                    //        break;
                    //    }
                    //}
                    //ViewData = IOitem;
                    //return true;
                    flg = true; break;
                }
                catch (Exception ex)
                { /*ProcessView.Add(new ListViewData() { RowData = "Open WISE Failed!", });*/
                    string str = ex.ToString();
                }
                
            }
            chiTxtbox.Text = id.ToString(); label22.Text = "5";
            if (flg) return true;

            return false;
        }
        private bool GetModbusDI_Val(int id)//20151016
        {
            try
            {   //Get modbus coil value.
                //var par = CustomController.ReadModbusCoils(DevInfo.DevModbus.CDI, DevInfo.DevModbus.LenCDI);
                var par = ModbusTCPService.ReadCoils(Device.MbCoils.DI, Device.MbCoils.lenDI);
                for (int i = 0; i < par.Length; i++)
                {
                    if (i == id - di_id_offset)
                    {
                        ProcessView(new ListViewData()
                        {
                            Step = "GetModbusDI_Val",
                            RowData = "Modbus = 0x" + (Device.MbCoils.DI + i).ToString("000")
                                + " ; Val: " + par[i].ToString()
                        });
                        return par[i];
                    }
                }
            }
            catch { }
            return false;
        }
        private bool GetMbCtOvf(int id)//20151110
        {
            try
            {   //Get modbus coil value.
                //var par = CustomController.ReadModbusCoils(DevInfo.DevModbus.CtOv, DevInfo.DevModbus.LenCtOv);
                var par = ModbusTCPService.ReadCoils(Device.MbCoils.CtOv, Device.MbCoils.lenCtOv);
                for (int i = 0; i < par.Length; i++)
                {
                    if (i == id - di_id_offset)
                    {
                        ProcessView(new ListViewData()
                        {
                            Step = "GetMbCtOvf",
                            RowData = "Modbus = 0x" + (Device.MbCoils.CtOv + i).ToString("000")
                                + " ; Val: " + par[i].ToString()
                        });
                        return par[i];
                    }
                }
            }
            catch { }
            return false;
        }
        private uint GetMbCtFrq_RegVal(int id)//20151110 new
        {
            uint res = 0;
            try
            {   //Get modbus coil value.
                //var par = CustomController.ReadModbusRegs(DevInfo.DevModbus.CtFq, DevInfo.DevModbus.LenCtFq);
                var par = ModbusTCPService.ReadHoldingRegs(Device.MbRegs.CtFq, Device.MbRegs.lenCtFq);
                for (int i = id * 2; i < par.Length; i++)
                {
                    ProcessView(new ListViewData()
                    {
                        Step = "GetMbCtFrq_RegVal",
                        RowData = "Modbus = 4x" + Device.MbRegs.CtFq.ToString("000")
                            + " ; Val: " + par[i].ToString()
                            + " ; Modbus = 4x" + (Device.MbRegs.CtFq + 1).ToString("000")
                            + " ; Val: " + par[i + 1].ToString()
                    });
                    //calculate modbus value to uint formate value
                    var big_edian = (uint)par[i + 1] << 16;
                    res = big_edian + (uint)par[i];
                    break;
                }
            }
            catch { }
            return res;
        }
        private bool InitRefCtrlDev()
        {
            //if (APAX_Demo) return true;
            //try
            //{
            //    CtrlDevInfo = CustomController.GetDeviceItemViewData(DevBase.APAX5070, true);//透過service更新資料
            //    if (CtrlDevInfo.MainDev.ModuleType != "")
            //    {
            //        return true;
            //    }
            //}
            //catch { ProcessView.Add(new ListViewData() { RowData = "Open Controller Device Failed!", }); }

            return true;
        }
        #region ----  Process Method  -----
        //IOItemViewData IOitem = new IOItemViewData();
        //IOItemViewData PACIOitem = new IOItemViewData();
        IOListData IOitem = new IOListData();
        private void ResultIni()
        {
            ResStr01 = ResStr02 = ResStr03 = ResStr04 = ResStr05 = ResStr06 = ResStr07 = "";
            SetData01 = SetData02 = SetData03 = SetData04 = SetData05 = SetData06 = SetData07 = 0;
            OutData01 = OutData02 = OutData03 = OutData04 = OutData05 = OutData06 = OutData07 = 0;
            MBData01 = MBData02 = MBData03 = MBData04 = MBData05 = MBData06 = MBData07 = "";
            DIStp = DIInvStp = DIL2HStp = DIH2LStp = DICtrStp = DICtrSTStp = 0;
            for (int i = 0; i < 7; i++) getTxtbox[i].Text = "";
        }
        int DIStp = 0;
        private void ChDI_input(int id)
        {
            if (DIStp >= 99)
            {
                ResStr01 = "Failed"; mainTaskStp++;
                FailCnt++;
            }
            else if (DIStp > 2 && DIStp < 99)
            {
                ResStr01 = "Passed"; mainTaskStp++;
            }
            else if (DIStp == 2)
            {
                //20151106
                bool MBres = GetModbusDI_Val(id);
                MBData01 = MBres.ToString();
                ProcessView(new ListViewData()
                {
                    Step = DIStp.ToString(),
                    RowData = "Val = " + IOitem.Val.ToString()
                        + " , Modbus = " + MBres.ToString()
                });

                if (IOitem.Val < 1 && !MBres) DIStp++;
                else if (IOitem.Val > 0 && MBres) DIStp++;
                else DIStp = 99;
            }
            else if (DIStp == 1)
            {
                GetDeviceItems(id);
                getTxtbox[0].Text = IOitem.Val.ToString();
                ProcessView(new ListViewData()
                {
                    Step = DIStp.ToString(),
                    RowData = "Val = " + IOitem.Val.ToString()
                        + " , SetData01 = " + SetData01.ToString()
                });
                if (SetData01 == 0 && IOitem.Val < 1) DIStp++;
                else if (SetData01 == 1 && IOitem.Val == 1) DIStp++;
                else DIStp = 99;
            }
            else if (DIStp == 0)
            {
                if (IOitem.Val > 0)
                {
                    OutData01 = OutputDO(false); SetData01 = 0;
                }
                else
                {
                    OutData01 = OutputDO(true); SetData01 = 1;
                }
                DIStp++;
            }
        }
        int DIInvStp = 0;
        private void ChDI_INV_Fun(int id)
        {
            if (DIInvStp >= 99)
            {
                ResStr02 = "Failed"; mainTaskStp++;
                FailCnt++;
            }
            else if (DIInvStp > 1 && DIInvStp < 99)
            {
                ResStr02 = "Passed"; mainTaskStp++;
            }
            else if (DIInvStp == 1)
            {
                GetDeviceItems(id);
                getTxtbox[1].Text = IOitem.Inv.ToString();
                //20151106
                bool MBres = GetModbusDI_Val(id);
                MBData02 = MBres.ToString();
                ProcessView(new ListViewData()
                {
                    Step = DIStp.ToString(),
                    RowData = "IO_Val = " + IOitem.Val.ToString()
                        + " , Modbus = " + MBres.ToString()
                });

                if (IOitem.Val > 0 && MBres)
                    DIInvStp++;
                else DIInvStp = 99;
            }
            else if (DIInvStp == 0)
            {
                IOitem.Inv = Ref_IO_Mod.Inv = SetData02 = 1;//enable
                HttpReqService.UpdateIOConfig(TransferIOModel());
                //CustomController.UpdateIOConfig(IOitem);
                //
                OutData02 = OutputDO(false);
                DIInvStp++;
            }
        }
        int DIL2HStp = 0;
        private void ChDI_LtoH_latch(int id)
        {
            if (DIL2HStp >= 99)
            {
                ResStr03 = "Failed"; mainTaskStp++;
                FailCnt++;
            }
            else if (DIL2HStp > 3 && DIL2HStp < 99)
            {
                ResStr03 = "Passed"; mainTaskStp++;
            }
            else if (DIL2HStp == 3)
            {
                GetDeviceItems(id);
                ProcessView(new ListViewData()
                {
                    Step = DIL2HStp.ToString(),
                    RowData = "IO_OvLch = " + IOitem.OvLch.ToString()
                });
                //20151106
                bool MBres = GetModbusDI_Val(id);
                MBData03 = MBres.ToString();

                if (IOitem.OvLch > 0 && MBres)
                    DIL2HStp++;
                else DIL2HStp = 99;
            }
            else if (DIL2HStp == 2)
            {
                GetDeviceItems(id);
                getTxtbox[2].Text = IOitem.Md.ToString();
                ProcessView(new ListViewData()
                {
                    Step = DIL2HStp.ToString(),
                    RowData = "IO_OvLch = " + IOitem.OvLch.ToString() +
                                " ,IO_Md = " + IOitem.Md.ToString()
                });
                if (IOitem.OvLch > 0 && IOitem.Md == 2)
                {
                    OutData03 = OutputDO(false);
                    DIL2HStp++;
                }
                else DIL2HStp = 99;
            }
            else if (DIL2HStp == 1)
            {
                //L --> H
                OutData03 = OutputDO(true);
                DIL2HStp++;
            }
            else if (DIL2HStp == 0)
            {
                Ref_IO_Mod.Md = IOitem.Md = SetData03 = 2;
                HttpReqService.UpdateIOConfig(TransferIOModel());
                //CustomController.UpdateIOConfig(IOitem);
                DIL2HStp++;
            }
        }
        int DIH2LStp = 0;
        private void ChDI_HtoL_latch(int id)
        {
            if (DIH2LStp >= 99)
            {
                ResStr04 = "Failed"; mainTaskStp++;
                FailCnt++;
            }
            else if (DIH2LStp > 4 && DIH2LStp < 99)
            {
                ResStr04 = "Passed"; mainTaskStp++;
            }
            else if (DIH2LStp == 4)
            {
                GetDeviceItems(id);
                ProcessView(new ListViewData()
                {
                    Step = DIH2LStp.ToString(),
                    RowData = "IO_OvLch = " + IOitem.OvLch.ToString()
                });
                //20151106
                bool MBres = GetModbusDI_Val(id);
                MBData04 = MBres.ToString();//modbus value would follow signal.

                if (IOitem.OvLch > 0 && !MBres)
                    DIH2LStp++;
                else DIH2LStp = 99;
            }
            else if (DIH2LStp == 3)
            {
                GetDeviceItems(id);
                getTxtbox[3].Text = IOitem.Md.ToString();
                ProcessView(new ListViewData()
                {
                    Step = DIH2LStp.ToString(),
                    RowData = "OvLch = " + IOitem.OvLch.ToString() +
                    " ,Md = " + IOitem.Md.ToString()
                });
                if (IOitem.OvLch > 0 && IOitem.Md == 3)
                {
                    OutData04 = OutputDO(true);
                    DIH2LStp++;
                }
                else DIH2LStp = 99;
            }
            else if (DIH2LStp == 2)
            {   //H --> L
                OutData04 = OutputDO(false);
                DIH2LStp++;
            }
            else if (DIH2LStp == 1)
            {
                Ref_IO_Mod.Md = IOitem.Md = SetData04 = 3;
                HttpReqService.UpdateIOConfig(TransferIOModel());
                //CustomController.UpdateIOConfig(IOitem);
                DIH2LStp++;
            }
            else if (DIH2LStp == 0)
            {
                ResetDefault();
                OutData04 = OutputDO(true);//Initial High
                DIH2LStp++;
            }
        }
        int DICtrStp = 0;
        private void ChDI_Ctr_Fun(int id)
        {
            if (DICtrStp >= 99)
            {
                ResStr05 = "Failed"; mainTaskStp++;
                FailCnt++;
            }
            else if (DICtrStp > 3 && DICtrStp < 99)
            {
                ResStr05 = "Passed"; mainTaskStp++;
            }
            else if (DICtrStp == 3)
            {
                GetDeviceItems(id);
                ProcessView(new ListViewData()
                {
                    Step = DICtrStp.ToString(),
                    RowData = "IO_Val = " + IOitem.Val.ToString()
                });
                uint mbVal = GetMbCtFrq_RegVal(id);
                MBData05 = mbVal.ToString();
                if (IOitem.Val >= 5 && IOitem.Md == 1 /*&& !APAX_Demo*/)
                    DICtrStp++;
                else DICtrStp = 99;
            }
            else if (DICtrStp == 2)
            {
                if (Controller_DO_PlsOut() || APAX_Demo) DICtrStp++;
                //ProcessStep = "APAX_DO_PlsOut outputing..." + DICtrStp.ToString();
            }
            else if (DICtrStp == 1)
            {
                GetDeviceItems(id);
                getTxtbox[4].Text = IOitem.Md.ToString();
                ProcessView(new ListViewData()
                {
                    Step = DICtrStp.ToString(),
                    RowData = "Md = " + IOitem.Md.ToString()
                });
                DICtrStp++;
            }
            else if (DICtrStp == 0)
            {
                Ref_IO_Mod.Md = IOitem.Md = SetData05 = 1;
                Ref_IO_Mod.CntIV = IOitem.CntIV = 0;
                HttpReqService.UpdateIOConfig(TransferIOModel());
                //CustomController.UpdateIOConfig(IOitem);
                Thread.Sleep(DelayT);
                //
                IOitem.ClrCnt = 1; IOitem.Cnting = 1;
                HttpReqService.UpdateIOValue(TransferIO_Val_Model());
                //CustomController.UpdateIOValue(IOitem);
                //
                OutputDO(false);
                //
                doPlCnt = 0; doPlsFlg = false;
                DICtrStp++;
            }
        }
        int DICtrSTStp = 0;
        private void ChDI_CtrST_Fun(int id)
        {
            uint StupVal = 4294967290;
            if (DICtrSTStp >= 99)
            {
                ResStr06 = "Failed"; mainTaskStp++;
                FailCnt++;
            }
            else if (DICtrSTStp > 7 && DICtrSTStp < 99)
            {
                ResStr06 = "Passed"; mainTaskStp++;
            }
            else if (DICtrSTStp == 7)
            {
                GetDeviceItems(id);
                //Get overflow flg
                var res = GetMbCtOvf(id);
                ProcessView(new ListViewData()
                {
                    Step = DICtrSTStp.ToString(),
                    RowData = "CtOv = " + res.ToString()
                                + " ; Val = " + IOitem.Val.ToString()
                });
                DICtrSTStp++;
            }
            else if (DICtrSTStp == 6)
            {
                if (Controller_DO_PlsOut() || APAX_Demo) DICtrSTStp++;
            }
            else if (DICtrSTStp == 5)
            {   //Initial DO
                OutputDO(false);
                doPlCnt = 0; doPlsFlg = false;
                DICtrSTStp++;
            }
            else if (DICtrSTStp == 4)
            {
                ProcessView(new ListViewData()
                {
                    Step = DICtrSTStp.ToString(),
                    RowData = "Val = " + IOitem.Val.ToString()
                });
                //
                uint mbVal = GetMbCtFrq_RegVal(id);
                MBData06 = mbVal.ToString();
                if (mbVal > StupVal) DICtrSTStp++;
                else DICtrSTStp = 99;
            }
            else if (DICtrSTStp == 3)
            {
                GetDeviceItems(id);
                ProcessView(new ListViewData()
                {
                    Step = DICtrSTStp.ToString(),
                    RowData = "Val = " + IOitem.Val.ToString()
                });
                if (IOitem.Val > StupVal && IOitem.Md == 1/*&& !APAX_Demo*/)
                    DICtrSTStp++;
                else DICtrSTStp = 99;
            }
            else if (DICtrSTStp == 2)
            {
                if (Controller_DO_PlsOut() || APAX_Demo) DICtrSTStp++;
                //ProcessStep = "APAX_DO_PlsOut outputing...";// +DICtrSTStp.ToString();
            }
            else if (DICtrSTStp == 1)
            {
                GetDeviceItems(id);
                getTxtbox[5].Text = IOitem.Md.ToString();
                ProcessView(new ListViewData()
                {
                    Step = DICtrSTStp.ToString(),
                    RowData = "Val = " + IOitem.Val.ToString()
                                + " ; Md = " + IOitem.Md.ToString()
                });
                DICtrSTStp++;
            }
            else if (DICtrSTStp == 0)
            {
                Ref_IO_Mod.Md = IOitem.Md = 1;
                Ref_IO_Mod.CntIV = IOitem.CntIV = StupVal;
                HttpReqService.UpdateIOConfig(TransferIOModel());
                //CustomController.UpdateIOConfig(IOitem);
                Thread.Sleep(DelayT);
                //
                IOitem.ClrCnt = 1; IOitem.Cnting = 1;
                HttpReqService.UpdateIOValue(TransferIO_Val_Model());
                //CustomController.UpdateIOValue(IOitem);
                //
                OutputDO(false);
                //
                doPlCnt = 0; doPlsFlg = false;
                DICtrSTStp++;
            }
        }
        private void ResetDefault()
        {
            IOitem.Md = 0; IOitem.Inv = 0;
            HttpReqService.UpdateIOConfig(TransferIOModel());
            //CustomController.UpdateIOConfig(IOitem);
            OutputDO(false);
        }
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
        #endregion

        //----------------------------------------------------------//
        void RefreshChTypeItem()
        {
            //APAX_Demo = false;
            //if (RefTypSelIdx == 6)//WISE-4051
            //{
            //    ChannelType = new ObservableCollection<string>
            //    {
            //        "All",
            //        "0","1","2","3","4","5","6","7",
            //    };
            //    Ref_IO_Mod = new WISE_IO_Model(24051);
            //    if (CtrlDevSelIdx == 0)
            //        OutDO_Slot_Idx = 9;
            //}
            //else if (RefTypSelIdx == 5)//WISE-4012
            //{
            //    ChannelType = new ObservableCollection<string>
            //    {
            //        "All",
            //        "0","1",
            //    };
            //    Ref_IO_Mod = new WISE_IO_Model(24012);
            //    if (CtrlDevSelIdx == 0)
            //        OutDO_Slot_Idx = 1;
            //}
            //else if (RefTypSelIdx == 4)//WISE-4060
            //{
            //    ChannelType = new ObservableCollection<string>
            //    {
            //        "All",
            //        "0","1","2","3",
            //    };
            //    Ref_IO_Mod = new WISE_IO_Model(24060);
            //    if (CtrlDevSelIdx == 0)
            //        OutDO_Slot_Idx = 5;
            //}
            //else if (RefTypSelIdx == 3)//WISE-4050
            //{
            //    ChannelType = new ObservableCollection<string>
            //    {
            //        "All",
            //        "0","1","2","3",
            //    };
            //    Ref_IO_Mod = new WISE_IO_Model(24050);
            //    if (CtrlDevSelIdx == 0)
            //        OutDO_Slot_Idx = 1;
            //}
            //else if (RefTypSelIdx == 2)//WISE-4060LAN
            //{
            //    ChannelType = new ObservableCollection<string>
            //    {
            //        "All",
            //        "0","1","2","3",
            //    };
            //    Ref_IO_Mod = new WISE_IO_Model(4060);
            //}
            //else if (RefTypSelIdx == 1)//WISE-4050LAN
            //{
            //    ChannelType = new ObservableCollection<string>
            //    {
            //        "All",
            //        "0","1","2","3",
            //    };
            //    Ref_IO_Mod = new WISE_IO_Model(4050);
            //}
            ////else if (RefTypSelIdx == 1)
            ////{
            ////    ChannelType = new ObservableCollection<string>
            ////    {
            ////        "All",
            ////        "0","1","2","3",
            ////    };
            ////    Ref_IO_Mod = new WISE_IO_Model(4010);
            ////}
            //else
            //{
            //    ChannelType = new ObservableCollection<string>
            //    {
            //        "All",
            //        "0","1","2","3",
            //    };
            //    APAX_Demo = true;
            //    Ref_IO_Mod = new WISE_IO_Model(4050);
            //}
        }
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

        private int OutputDO(bool TF)
        {
            if (APAX_Demo) return 0;

            //if (CtrlDevSelIdx == 0)
                OutputAPAX5070(TF);
            //else
                //OutputPAC(TF);

            return TF ? 1 : 0;
        }
        private void OutputAPAX5070(bool TF)
        {
            int val = TF ? (int)1 : 0;
            int idx = OutDO_Slot_Idx + Ref_IO_Mod.Ch;
            APAX5070Service.ForceSigCoil(idx, val);
            Thread.Sleep(100);
        }

        bool doPlsFlg;

        

        int doPlCnt = 0;
        private bool Controller_DO_PlsOut()
        {
            ProcessStep = "Ctl_DO_PlsOut outputing..." + doPlCnt.ToString();
            if (doPlCnt > 4)
            {
                return true;
            }
            else if (doPlsFlg)
            {
                doPlsFlg = OutputDO(false) > 0 ? true : false;
                doPlCnt++;
            }
            else
            {
                doPlsFlg = OutputDO(true) > 0 ? true : false;
            }
            return false;
        }


    }//class
    public class ListViewData
    {
        public int Ch { get; set; }
        public string Step { get; set; }
        public string Result { get; set; }
        public string RowData { get; set; }
    }
    public enum WISE_AT_DI_Task
    {
        wCh_Init = 1,
        wCh_DI_Fun = 2,
        wCh_DI_Fun_res = 3,
        wCh_DI_Inv = 4,
        wCh_DI_Inv_res = 5,
        wCh_DI_RstToInit = 6,
        wCh_DI_LtoH_Latch = 7,
        wCh_DI_LtoH_Latch_res = 8,
        wCh_DI_HtoL_Latch = 9,
        wCh_DI_HtoL_Latch_res = 10,
        wCh_DI_Counter = 11,
        wCh_DI_Counter_res = 12,
        wCh_DI_Counter_Stup = 13,
        wCh_DI_Counter_Stup_res = 14,

        wFinished = 15,


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
