namespace Model
{
    public class WISE_IO_Model : IOListData
    {
        public string ModelType;
        public int DI_num;
        public int DO_num;
        public int AI_num;
        public int AO_num;
        public int Cnt_num;
        public int APAX_AO_HiVal;
        public int APAX_AO_MdVal;
        public int APAX_AO_LoVal;

        public int[] AIRng;

        /// <summary>
        /// main
        /// </summary>
        public WISE_IO_Model(int _typ)
        {
            if (_typ == (int)WISEType.WISE4010LAN)
            {
                ModelType = WISEType.WISE4010LAN.ToString();
                AI_num = 4; DO_num = 4;
                AIRng = new int[2];
                AIRng[0] = (int)ValueRange.mA_4To20; AIRng[1] = (int)ValueRange.mA_0To20;
                APAX_AO_HiVal = 20000;
                APAX_AO_MdVal = 10000;
                APAX_AO_LoVal = 4000;
            }
            else if (_typ == (int)WISEType.WISE4012)
            {
                ModelType = WISEType.WISE4012.ToString();
                AI_num = 4; DO_num = 2; DI_num = 4;
                AIRng = new int[13];
                AIRng[0] = (int)ValueRange.mV_0To150; AIRng[1] = (int)ValueRange.mV_0To500; AIRng[2] = (int)ValueRange.V_0To1;
                AIRng[3] = (int)ValueRange.V_0To5; AIRng[4] = (int)ValueRange.V_0To10; AIRng[5] = (int)ValueRange.mV_Neg150To150;
                AIRng[6] = (int)ValueRange.mV_Neg500To500; AIRng[7] = (int)ValueRange.V_Neg1To1; AIRng[8] = (int)ValueRange.V_Neg5To5;
                AIRng[9] = (int)ValueRange.V_Neg10To10; AIRng[10] = (int)ValueRange.mA_4To20; AIRng[11] = (int)ValueRange.mA_0To20;
                AIRng[12] = (int)ValueRange.mA_Neg20To20;
                APAX_AO_HiVal = 10000;
                APAX_AO_MdVal = 5000;
                APAX_AO_LoVal = 0;
            }
            else if (_typ == (int)WISEType.WISE4050 || _typ == (int)WISEType.WISE4050LAN)
            {
                ModelType = WISEType.WISE4050.ToString();
                DI_num = 4; DO_num = 4;
            }
            else if (_typ == (int)WISEType.WISE4060 || _typ == (int)WISEType.WISE4060LAN)
            {
                ModelType = WISEType.WISE4060.ToString();
                DI_num = 4; DO_num = 4;
            }
            else if (_typ == (int)WISEType.WISE4051)
            {
                ModelType = WISEType.WISE4051.ToString();
                DI_num = 8;
            }


            
        }

    }
}


namespace Model
{
    public class IOBasicModel
    {
        
    }
}