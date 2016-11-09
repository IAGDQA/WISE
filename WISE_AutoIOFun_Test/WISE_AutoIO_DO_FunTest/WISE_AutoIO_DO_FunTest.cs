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

namespace WISE_AutoIO_DO_FunTest
{
    public partial class WISE_AutoIO_DO_FunTest : Form, iATester.iCom
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
        int do_id_offset = 20;
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

        //======================================================//
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
        int _ref_do_num = 0;
        public string OutDO_Ch_Len
        {
            get
            { return "0 ~ " + (_ref_do_num - 1).ToString(); }
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
        public WISE_AutoIO_DO_FunTest()
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

            SelAllChkbox.Checked = StpChkIdx0 = true;//select all steps
            //ProcessView = new ObservableCollection<ListViewData>();
        }

        private void WISE_AutoIO_DO_FunTest_Load(object sender, EventArgs e)
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
            GetParaFromFile();
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
                case (int)WISE_AT_DO_Task.wCh_Init:
                    //ProcessView.Add(new ListViewData() { Ch = Ref_IO_Mod.Ch, Step = "Initialized", Result = "" });
                    ProcessView(new ListViewData() { Ch = Ref_IO_Mod.Ch, Step = "Initialized", Result = "" });
                    GetDeviceItems(do_id_offset + Ref_IO_Mod.Ch);
                    mainTaskStp++;
                    break;
                case (int)WISE_AT_DO_Task.wCh_DO_Fun:
                    if (StpChkIdx1)
                    {
                        ProcessStep = WISE_AT_DO_Task.wCh_DO_Fun.ToString();
                        ResStr01 = ChDO_output(do_id_offset + Ref_IO_Mod.Ch) ? "Passed" : "Failed";
                        ProcessView(new ListViewData() { Ch = Ref_IO_Mod.Ch, Step = ProcessStep, Result = ResStr01 });
                    }
                    else mainTaskStp++;
                    break;
                case (int)WISE_AT_DO_Task.wCh_DO_PulseOutMd0:
                    if (StpChkIdx2)
                    {
                        ProcessStep = WISE_AT_DO_Task.wCh_DO_PulseOutMd0.ToString();
                        ChDO_PlsOutMd0(do_id_offset + Ref_IO_Mod.Ch);
                    }
                    else mainTaskStp = (int)WISE_AT_DO_Task.wCh_DO_PulseOutMd1;
                    break;
                case (int)WISE_AT_DO_Task.wCh_DO_PulseOutMd0_res:
                    ProcessView(new ListViewData() { Ch = Ref_IO_Mod.Ch, Step = ProcessStep, Result = ResStr02 });
                    mainTaskStp++;
                    break;
                case (int)WISE_AT_DO_Task.wCh_DO_PulseOutMd1:
                    if (StpChkIdx3)
                    {
                        ProcessStep = WISE_AT_DO_Task.wCh_DO_PulseOutMd1.ToString();
                        ChDO_PlsOutMd1(do_id_offset + Ref_IO_Mod.Ch);
                    }
                    else mainTaskStp = (int)WISE_AT_DO_Task.wCh_DO_LtoH_Delay;
                    break;
                case (int)WISE_AT_DO_Task.wCh_DO_PulseOutMd1_res:
                    ProcessView(new ListViewData() { Ch = Ref_IO_Mod.Ch, Step = ProcessStep, Result = ResStr03 });
                    mainTaskStp++;
                    break;
                case (int)WISE_AT_DO_Task.wCh_DO_RstDO:
                    ProcessStep = WISE_AT_DO_Task.wCh_DO_RstDO.ToString();
                    ResetDO(); mainTaskStp++;
                    break;
                case (int)WISE_AT_DO_Task.wCh_DO_LtoH_Delay:
                    if (StpChkIdx4)
                    {
                        ProcessStep = WISE_AT_DO_Task.wCh_DO_LtoH_Delay.ToString();
                        ChDO_LtoH_delay(do_id_offset + Ref_IO_Mod.Ch);
                    }
                    else mainTaskStp = (int)WISE_AT_DO_Task.wCh_DO_HtoL_Delay;
                    break;
                case (int)WISE_AT_DO_Task.wCh_DO_LtoH_Delay_res:
                    ProcessStep = WISE_AT_DO_Task.wCh_DO_LtoH_Delay_res.ToString();
                    ChDO_LtoH_delay_res(do_id_offset + Ref_IO_Mod.Ch);
                    break;
                case (int)WISE_AT_DO_Task.wCh_DO_HtoL_Delay:
                    if (StpChkIdx5)
                    {
                        ProcessStep = WISE_AT_DO_Task.wCh_DO_HtoL_Delay.ToString();
                        ChDO_HtoL_delay(do_id_offset + Ref_IO_Mod.Ch);
                    }
                    else mainTaskStp = (int)WISE_AT_DO_Task.wCh_DO_ALDrived;
                    break;
                case (int)WISE_AT_DO_Task.wCh_DO_HtoL_Delay_res:
                    ProcessStep = WISE_AT_DO_Task.wCh_DO_HtoL_Delay_res.ToString();
                    ChDO_HtoL_delay_res(do_id_offset + Ref_IO_Mod.Ch);
                    break;
                case (int)WISE_AT_DO_Task.wCh_DO_ALDrived:
                    if (StpChkIdx6)
                    {
                        ProcessStep = WISE_AT_DO_Task.wCh_DO_ALDrived.ToString(); mainTaskStp++;
                    }
                    else mainTaskStp++;
                    break;


                case (int)WISE_AT_DO_Task.wFinished:
                    ProcessStep = WISE_AT_DO_Task.wFinished.ToString();
                    ProcessView(new ListViewData() { Ch = Ref_IO_Mod.Ch, Step = "Finished", Result = "" });

                    if (ChSelIdx == 0 && Ref_IO_Mod.Ch < (Ref_IO_Mod.DO_num - 1))
                    {
                        Ref_IO_Mod.Ch++; GetDeviceItems(do_id_offset + Ref_IO_Mod.Ch); VerInit();
                    }
                    else
                    {
                        timer.Stop(); RunBtnStr = "Run";
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
            if (Device.IPAddress == "" || Device.IPAddress == null) Device.IPAddress = tempIP;
            SetParaToFile();
            if (HttpReqService.HttpReqTCP_Connet(Device))
            {
                Device = HttpReqService.GetDevice();
                ConnReadyFlg = true;
                typTxtbox.Text = Device.ModuleType;
                tssLabel1.Text = "";
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
            if (Device.ModuleType == "WISE-4051")
            {
                Ref_IO_Mod = new WISE_IO_Model(24051);
            }
            else if (Device.ModuleType == "WISE-4012")
            {
                Ref_IO_Mod = new WISE_IO_Model(24012);
            }
            else if (Device.ModuleType == "WISE-4012E")
            {
                Ref_IO_Mod = new WISE_IO_Model(34012);
            }
            else if (Device.ModuleType == "WISE-4060")
            {
                Ref_IO_Mod = new WISE_IO_Model(4060);
            }
            else if (Device.ModuleType == "WISE-4050")
            {
                Ref_IO_Mod = new WISE_IO_Model(4050);
            }
            else if (Device.ModuleType == "WISE-4060/LAN")
            {
                Ref_IO_Mod = new WISE_IO_Model(4060);
            }
            else if (Device.ModuleType == "WISE-4050/LAN")
            {
                Ref_IO_Mod = new WISE_IO_Model(4050);
            }
            else
            {
                APAX_Demo = true;
                Ref_IO_Mod = new WISE_IO_Model(24051);
            }

            _ref_do_num = Ref_IO_Mod.DO_num;
            OutDO_Ch_Len = ""; lenTxtbox.Text = _ref_do_num.ToString();
            //
            if (!GetDeviceItems(do_id_offset + Ref_IO_Mod.Ch)
                || !InitRefCtrlDev()) return false;
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
            TestResult = "N/A";
            mainTaskStp = 0;
            ResultIni();
        }
        private bool GetDeviceItems(int id)
        {
            bool flg = false; label22.Text = "GetDeviceItems.........";
            while (true)
            {
                try
                {
                    List<IOListData> IO_Data = (List<IOListData>)HttpReqService.GetListOfIOItems("");
                    foreach (var item in IO_Data)
                    {
                        if (item.Id == id && item.Ch == id - do_id_offset)
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
                    if(IOitem.Id!=null)
                    {
                        flg = true; break;
                    }
                    
                }
                catch (Exception ex)
                { /*ProcessView.Add(new ListViewData() { RowData = "Open WISE Failed!", });*/
                    string str = ex.ToString();
                }
            }
            chiTxtbox.Text = (id- do_id_offset).ToString(); label22.Text = "";
            if (flg) return true;

            return false;
        }
        private bool GetModbusDO_Val(int id)//20151016
        {
            try
            {   //Get modbus coil value.
                var par = ModbusTCPService.ReadCoils(Device.MbCoils.DO, Device.MbCoils.lenDO);
                for (int i = 0; i < par.Length; i++)
                {
                    if (i == id - do_id_offset)
                    {
                        ProcessView(new ListViewData()
                        {
                            Step = "GetModbusDO_Val",
                            RowData = "Modbus = 0x" + (Device.MbCoils.DO + i).ToString("000")
                                + " ; Val: " + par[i].ToString()
                        });
                        return par[i];
                    }
                }
            }
            catch { }
            return false;
        }
        private uint GetMbPsAV_RegVal(int id)//20151110 new
        {
            uint res = 0;
            try
            {   //Get modbus reg value.
                var par = ModbusTCPService.ReadHoldingRegs(Device.MbRegs.PsAV, Device.MbRegs.lenPsAV);

                for (int i = (id - do_id_offset) * 2; i < par.Length; i++)
                {
                    ProcessView(new ListViewData()
                    {
                        Step = "GetMbPsAV_RegVal",
                        RowData = "Modbus = 4x" + Device.MbRegs.PsAV.ToString("000")
                                + " ; Val: " + par[i].ToString()
                                + " ; Modbus = 4x" + (Device.MbRegs.PsAV + 1).ToString("000")
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

        //================================================================================//
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
            ChDO_PlsMd0Stp = ChDO_PlsMd1Stp = 0;
            for (int i = 0; i < 7; i++) getTxtbox[i].Text = "";
        }

        private bool ChDO_output(int id)
        {
            mainTaskStp++;
            Ref_IO_Mod.Val = (uint)1; SetData01 = 1;
            ResetDO();
            Thread.Sleep(DelayT);
            GetDeviceItems(id);
            getTxtbox[0].Text = IOitem.Val.ToString();
            ProcessView(new ListViewData() { RowData = "IO_Val = " + IOitem.Val.ToString() });
            //20151016
            bool MBres = GetModbusDO_Val(id);
            MBData01 = MBres.ToString();

            if (IOitem.Val > 0 && MBres)
                return true;

            return false;
        }
        int ChDO_PlsMd0Stp = 0;
        private void ChDO_PlsOutMd0(int id)
        {
            if (ChDO_PlsMd0Stp >= 99)
            {
                ResStr02 = "Failed"; mainTaskStp++;
            }
            else if (ChDO_PlsMd0Stp > 2 && ChDO_PlsMd0Stp < 99)
            {
                ResStr02 = "Passed"; mainTaskStp++;
            }
            else if (ChDO_PlsMd0Stp == 2)
            {
                ProcessView(new ListViewData()
                {
                    RowData = "Md = " + IOitem.Md.ToString()
                                + "Val = " + IOitem.Val.ToString()
                });
                //
                uint mbVal = GetMbPsAV_RegVal(id);
                MBData02 = mbVal.ToString();
                if (mbVal > 4294967290) ChDO_PlsMd0Stp = 99;

                ChDO_PlsMd0Stp++;
            }
            else if (ChDO_PlsMd0Stp == 1)
            {
                GetDeviceItems(id);
                getTxtbox[1].Text = IOitem.Val.ToString();
                if (IOitem.Val < 4294967290 && IOitem.Val > 65535) ChDO_PlsMd0Stp++;
                //else ChDO_PlsMd0Stp = 0;
                ProcessStep = "DO_PlsOut outputing..." + IOitem.Val.ToString();
            }
            else if (ChDO_PlsMd0Stp == 0)
            {
                Ref_IO_Mod.Md = IOitem.Md = SetData02 = 1;
                IOitem.PsLo = 5000; IOitem.PsHi = 6000;
                HttpReqService.UpdateIOConfig(TransferIOModel());
                //CustomController.UpdateIOConfig(IOitem);
                Thread.Sleep(DelayT);
                //
                IOitem.PsCtn = 0; IOitem.Val = 4294967295;
                HttpReqService.UpdateIOValue(TransferIO_Val_Model());
                //CustomController.UpdateIOValue(IOitem);
                ChDO_PlsMd0Stp++;
            }
        }
        int ChDO_PlsMd1Stp = 0;
        private void ChDO_PlsOutMd1(int id)//Continue mode problem in 20151013 fixed
        {
            if (ChDO_PlsMd1Stp >= 99)
            {
                ResStr03 = "Failed"; mainTaskStp++;
            }
            else if (ChDO_PlsMd1Stp > 1 && ChDO_PlsMd1Stp < 99)
            {
                ResStr03 = "Passed"; mainTaskStp++;
            }
            //Modbus address no define continue value
            //else if (ChDO_PlsMd1Stp == 2)
            //{
            //    ProcessView.Add(new ListViewData()
            //    {
            //        RowData = "Val = " + IOitem.Val.ToString()
            //    });
            //    uint mbVal = GetMbPsAV_RegVal(id);
            //    MBData03 = mbVal.ToString();
            //    if (mbVal > 4294967290) ChDO_PlsMd1Stp++;
            //    else ChDO_PlsMd1Stp = 99;
            //}
            else if (ChDO_PlsMd1Stp == 1)
            {
                GetDeviceItems(id);
                getTxtbox[2].Text = IOitem.Val.ToString();
                if (IOitem.Val > 4294967290) ChDO_PlsMd1Stp++;
                else ChDO_PlsMd1Stp = 0;
                ProcessStep = "DO_PlsOut outputing..." + IOitem.Val.ToString();
            }
            else if (ChDO_PlsMd1Stp == 0)
            {
                Ref_IO_Mod.Md = IOitem.Md = SetData03 = 1;
                IOitem.PsLo = 5000; IOitem.PsHi = 6000;
                HttpReqService.UpdateIOConfig(TransferIOModel());
                //CustomController.UpdateIOConfig(IOitem);
                //
                IOitem.PsCtn = 1; //IOitem.Val = 0;
                HttpReqService.UpdateIOValue(TransferIO_Val_Model());
                //CustomController.UpdateIOValue(IOitem);
                ChDO_PlsMd1Stp++;
            }
        }

        private bool ChDO_LtoH_delay(int id)
        {
            mainTaskStp++;
            Ref_IO_Mod.Md = IOitem.Md = SetData04 = 2;
            Ref_IO_Mod.LDT = IOitem.LDT = 65535;
            Ref_IO_Mod.HDT = IOitem.HDT = 0;
            HttpReqService.UpdateIOConfig(TransferIOModel());
            //CustomController.UpdateIOConfig(IOitem);
            Thread.Sleep(500);
            //
            IOitem.Val = 1;//Let DO initial poll-high.
            HttpReqService.UpdateIOValue(TransferIO_Val_Model());
            //CustomController.UpdateIOValue(IOitem);
            ProcessView(new ListViewData() { RowData = "LDT = " + IOitem.LDT.ToString() });

            return true;
        }
        bool STflg = false; int waitCnt = 0;
        private bool ChDO_LtoH_delay_res(int id)
        {
            GetDeviceItems(id);
            getTxtbox[3].Text = IOitem.Val.ToString();
            if (STflg)
            {
                if (IOitem.Val == 1)
                {
                    ProcessView(new ListViewData()
                    {
                        Ch = Ref_IO_Mod.Ch,
                        Step = ProcessStep,
                        Result = ResStr04,
                        RowData = "Cnt = " + waitCnt.ToString()
                    });
                    //20151016
                    bool MBres = GetModbusDO_Val(id);
                    MBData04 = MBres.ToString();

                    if (MBres) ResStr04 = "Passed";
                    else ResStr04 = "Failed";

                    mainTaskStp++; STflg = false; waitCnt = 0;
                }
                else if (waitCnt > 10)//10 == 10 sec
                {
                    ResStr04 = "Failed";
                    ProcessView(new ListViewData()
                    {
                        Ch = Ref_IO_Mod.Ch,
                        Step = ProcessStep,
                        Result = ResStr04,
                        RowData = "Cnt = " + waitCnt.ToString() + " Sec"
                    });
                    mainTaskStp++; STflg = false; waitCnt = 0;
                }
                waitCnt++;
            }
            else if (!STflg && IOitem.Val == 1)//Initial get val is 1.
            {
                ResStr04 = "Failed";
                ProcessView(new ListViewData()
                {
                    Ch = Ref_IO_Mod.Ch,
                    Step = ProcessStep,
                    Result = ResStr04,
                    RowData = "Val = " + IOitem.Val.ToString()
                });
                mainTaskStp++;
            }
            else STflg = true;

            return true;
        }

        private bool ChDO_HtoL_delay(int id)
        {
            mainTaskStp++;
            Ref_IO_Mod.Md = IOitem.Md = SetData05 = 3;
            Ref_IO_Mod.HDT = IOitem.HDT = 65535;
            Ref_IO_Mod.LDT = IOitem.LDT = 0;
            HttpReqService.UpdateIOConfig(TransferIOModel());
            //CustomController.UpdateIOConfig(IOitem);
            Thread.Sleep(500);
            GetDeviceItems(id);
            //
            if (IOitem.Val == 0)
            {
                IOitem.Val = 1;
                HttpReqService.UpdateIOValue(TransferIO_Val_Model());
                //CustomController.UpdateIOValue(IOitem);
                Thread.Sleep(500);
            }
            //
            IOitem.Val = 0;
            HttpReqService.UpdateIOValue(TransferIO_Val_Model());
            //CustomController.UpdateIOValue(IOitem);
            ProcessView(new ListViewData() { RowData = "HDT = " + IOitem.HDT.ToString() });

            return true;
        }

        private bool ChDO_HtoL_delay_res(int id)
        {
            GetDeviceItems(id);
            getTxtbox[4].Text = IOitem.Val.ToString();
            if (STflg)
            {
                if (IOitem.Val == 0)
                {
                    ProcessView(new ListViewData()
                    {
                        Ch = Ref_IO_Mod.Ch,
                        Step = ProcessStep,
                        Result = ResStr05,
                        RowData = "Cnt = " + waitCnt.ToString()
                    });
                    //20151016
                    bool MBres = GetModbusDO_Val(id);
                    MBData05 = MBres.ToString();

                    if (!MBres) ResStr05 = "Passed";//H->L
                    else ResStr05 = "Failed";

                    mainTaskStp++; STflg = false; waitCnt = 0;
                }
                else if (waitCnt > 10)//10 == 10 sec
                {
                    ResStr05 = "Failed";
                    ProcessView(new ListViewData()
                    {
                        Ch = Ref_IO_Mod.Ch,
                        Step = ProcessStep,
                        Result = ResStr05,
                        RowData = "Cnt = " + waitCnt.ToString() + " Sec"
                    });
                    mainTaskStp++; STflg = false; waitCnt = 0;
                }
                waitCnt++;
            }
            else if (!STflg && IOitem.Val == 0)//Initial get val is 0.
            {
                ResStr05 = "Failed";
                ProcessView(new ListViewData()
                {
                    Ch = Ref_IO_Mod.Ch,
                    Step = ProcessStep,
                    Result = ResStr05,
                    RowData = "Val = " + IOitem.Val.ToString()
                });
                mainTaskStp++;
            }
            else STflg = true;

            return true;
        }
        
        private void ResetDO()
        {
            IOitem.Md = 0;
            HttpReqService.UpdateIOConfig(TransferIOModel());
            Thread.Sleep(500);
            IOitem.Val = 1;
            //CustomController.UpdateIOValue(IOitem);
            HttpReqService.UpdateIOValue(TransferIO_Val_Model());
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

        int doPlCnt = 0; bool doPlsFlg;
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
    public enum WISE_AT_DO_Task
    {
        wCh_Init = 1,
        wCh_DO_Fun = 2,
        wCh_DO_PulseOutMd0 = 3,
        wCh_DO_PulseOutMd0_res = 4,
        wCh_DO_PulseOutMd1 = 5,
        wCh_DO_PulseOutMd1_res = 6,
        wCh_DO_RstDO = 7,
        wCh_DO_LtoH_Delay = 8,
        wCh_DO_LtoH_Delay_res = 9,
        wCh_DO_HtoL_Delay = 10,
        wCh_DO_HtoL_Delay_res = 11,
        wCh_DO_ALDrived = 12,//寫在AI測試裡面

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
