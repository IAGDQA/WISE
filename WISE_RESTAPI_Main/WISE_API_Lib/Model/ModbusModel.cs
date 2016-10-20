namespace Model
{
    /// <summary>
    /// 1.需要連結MainController-> GetDeviceItemViewData() 做更新
    /// 2.同時更新DeviceViewData struct
    /// </summary>
    public class ModbusCoilModel
    {
        //Base Config
        public int Md
        {
            get;
            set;
        }
        public int DI
        {
            get;
            set;
        }
        public int CtS
        {
            get;
            set;
        }
        public int CtClr
        {
            get;
            set;
        }
        public int CtOv
        {
            get;
            set;
        }
        public int Lch
        {
            get;
            set;
        }
        public int DO
        {
            get;
            set;
        }
        public int AIHR
        {
            get;
            set;
        }
        public int AILR
        {
            get;
            set;
        }
        public int AIB
        {
            get;
            set;
        }
        public int HAlm
        {
            get;
            set;
        }
        public int LAlm
        {
            get;
            set;
        }
        public int GCtClr
        {
            get;
            set;
        }
        //Base Length
        public int lenDI
        {
            get;
            set;
        }
        public int lenCtS
        {
            get;
            set;
        }
        public int lenCtClr
        {
            get;
            set;
        }
        public int lenCtOv
        {
            get;
            set;
        }
        public int lenLch
        {
            get;
            set;
        }
        public int lenDO
        {
            get;
            set;
        }
        public int lenAIHR
        {
            get;
            set;
        }
        public int lenAILR
        {
            get;
            set;
        }
        public int lenAIB
        {
            get;
            set;
        }
        public int lenHAlm
        {
            get;
            set;
        }
        public int lenLAlm
        {
            get;
            set;
        }
        public int lenGCtClr
        {
            get;
            set;
        }
    }

    public class ModbusRegModel
    {
        //Base Config
        public int Md
        {
            get;
            set;
        }
        public int DI
        {
            get;
            set;
        }
        public int CtFq
        {
            get;
            set;
        }
        public int DO
        {
            get;
            set;
        }
        public int PsLo
        {
            get;
            set;
        }
        public int PsHi
        {
            get;
            set;
        }
        public int PsAV
        {
            get;
            set;
        }
        public int PsIV
        {
            get;
            set;
        }
        public int AI
        {
            get;
            set;
        }
        public int AIFl
        {
            get;
            set;
        }
        public int AICd
        {
            get;
            set;
        }
        public int AlCh
        {
            get;
            set;
        }
        public int AlSc
        {
            get;
            set;
        }
        public int HisH
        {
            get;
            set;
        }
        public int HisL
        {
            get;
            set;
        }
        public int AIF
        {
            get;
            set;
        }
        public int HisHF
        {
            get;
            set;
        }
        public int HisLF
        {
            get;
            set;
        }
        public int AICh
        {
            get;
            set;
        }
        public int AISc
        {
            get;
            set;
        }
        public int AO
        {
            get;
            set;
        }
        public int AOFl
        {
            get;
            set;
        }
        public int AOCd
        {
            get;
            set;
        }
        public int Slew
        {
            get;
            set;
        }
        public int AOSu
        {
            get;
            set;
        }
        public int AOSe
        {
            get;
            set;
        }
        public int AODi
        {
            get;
            set;
        }
        public int GCLCt
        {
            get;
            set;
        }
        public int GCLFl
        {
            get;
            set;
        }
        public int MNm
        {
            get;
            set;
        }

        //Base Length
        public int lenDI
        {
            get;
            set;
        }
        public int lenCtFq
        {
            get;
            set;
        }
        public int lenDO
        {
            get;
            set;
        }
        public int lenPsLo
        {
            get;
            set;
        }
        public int lenPsHi
        {
            get;
            set;
        }
        public int lenPsAV
        {
            get;
            set;
        }
        public int lenPsIV
        {
            get;
            set;
        }
        public int lenAI
        {
            get;
            set;
        }
        public int lenAIFl
        {
            get;
            set;
        }
        public int lenAICd
        {
            get;
            set;
        }
        public int lenAlCh
        {
            get;
            set;
        }
        public int lenAlSc
        {
            get;
            set;
        }
        public int lenHisH
        {
            get;
            set;
        }
        public int lenHisL
        {
            get;
            set;
        }
        public int lenAIF
        {
            get;
            set;
        }
        public int lenHisHF
        {
            get;
            set;
        }
        public int lenHisLF
        {
            get;
            set;
        }
        public int lenAICh
        {
            get;
            set;
        }
        public int lenAISc
        {
            get;
            set;
        }
        public int lenAO
        {
            get;
            set;
        }
        public int lenAOFl
        {
            get;
            set;
        }
        public int lenAOCd
        {
            get;
            set;
        }
        public int lenSlew
        {
            get;
            set;
        }
        public int lenAOSu
        {
            get;
            set;
        }
        public int lenAOSe
        {
            get;
            set;
        }
        public int lenAODi
        {
            get;
            set;
        }
        public int lenGCLCt
        {
            get;
            set;
        }
        public int lenGCLFl
        {
            get;
            set;
        }
        public int lenMNm
        {
            get;
            set;
        }
    }
}
