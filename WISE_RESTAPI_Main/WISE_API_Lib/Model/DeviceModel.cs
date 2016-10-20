namespace Model
{
    /// <summary>
    /// 1.需要連結MainController-> GetDeviceItemViewData() 做更新
    /// 2.同時更新DeviceViewData struct
    /// </summary>
    public class DeviceModel : HttpReqModel
    {
        //Modbus
        public ModbusCoilModel MbCoils
        {
            get;
            set;
        }

        public ModbusRegModel MbRegs
        {
            get;
            set;
        }        

        //Common
        public string IPAddress
        {
            get;
            set;
        }
        public string ModuleType
        {
            get;
            set;
        }
        public string FirmwareVer
        {
            get;
            set;
        }
        public string Account
        {
            get;
            set;
        }
        public string Password
        {
            get;
            set;
        }

        public int Port
        {
            get;
            set;
        }

        public int SlotNum
        {
            get;
            set;
        }

        public int ChannelNum
        {
            get;
            set;
        }

        public int DiTotal
        {
            get;
            set;
        }

        public int DoTotal
        {
            get;
            set;
        }

        public int RLaTotal
        {
            get;
            set;
        }

        public int AiTotal
        {
            get;
            set;
        }

        public int UInTotal
        {
            get;
            set;
        }

        public int AoTotal
        {
            get;
            set;
        }

        public int CntTotal
        {
            get;
            set;
        }

        public bool ConnFlg
        {
            get;
            set;
        }

        public string Exception
        {
            get;
            set;
        }
        //AI
        public int Res
        {
            get;
            set;
        }
        public int EnB
        {
            get;
            set;
        }
        public int BMd
        {
            get;
            set;
        }
        public int AiT
        {
            get;
            set;
        }
        public int Smp
        {
            get;
            set;
        }
        public int AvgM
        {
            get;
            set;
        }
        //
        public int ModbusTimeOut
        {
            get;
            set;
        }
        public int ModbusAddr
        {
            get;
            set;
        }
    }

    public class IOModel : IOListData
    {
        
    }
}
