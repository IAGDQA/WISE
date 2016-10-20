namespace Model
{
    /// <summary>
    /// Summary information for a Customer
    /// As a 'cut down' version of Customer information, this class is used
    /// for lists of Customers, for example, to avoid having to get a complete
    /// Customer object
    /// </summary>
    public class IOListData
    {
        /// <summary>
        /// The unique Id assigned to this Customer in the Data Store
        /// </summary>
        #region -- Basic --
        public int? Id
        {
            get;
            set;
        }

        public int Ch
        {
            get;
            set;
        }

        public uint Val //fix DO/DI read data 20151013
        {
            get;
            set;
        }

        public string Tag
        {
            get;
            set;
        }

        public string Exception
        {
            get;
            set;
        }
        
        #endregion
        //AI
        #region -- AI --
        
        public int En  //AI read
        {
            get;
            set;
        }
        public int Rng //AI read
        {
            get;
            set;
        }
        public int Evt  //AI read
        {
            get;
            set;
        }
        public int Eg //AI read
        {
            get;
            set;
        }
        public string Val_Eg  //AI
        {
            get;
            set;
        }

        public string Val_Hex  //AI
        {
            get;
            set;
        }

        public string Val_Dec //AI
        {
            get;
            set;
        }
        public int LoA //AI read
        {
            get;
            set;
        }
        public int HiA //AI read
        {
            get;
            set;
        }

        public float EgF //AI read
        {
            get;
            set;
        }


        //config
        public int cEn  //AI
        {
            get;
            set;
        }
        public int cRng
        {
            get;
            set;
        }
        public int? EnLA
        {
            get;
            set;
        }
        public int? EnHA
        {
            get;
            set;
        }
        public int? LAMd
        {
            get;
            set;
        }
        public int? HAMd
        {
            get;
            set;
        }
        public string cLoA
        {
            get;
            set;
        }
        public string cHiA
        {
            get;
            set;
        }
        public string LoS
        {
            get;
            set;
        }
        public string HiS
        {
            get;
            set;
        }
        public int? Res
        {
            get;
            set;
        }
        public int? EnB
        {
            get;
            set;
        }
        public int? BMd
        {
            get;
            set;
        }
        public int? AiT
        {
            get;
            set;
        }
        public int? Smp
        {
            get;
            set;
        }
        public int? AvgM
        {
            get;
            set;
        }

        #endregion

        //DI/O
        #region -- DI --
        public int? Md
        {
            get;
            set;
        }
        public int? Inv
        {
            get;
            set;
        }
        public int? Fltr
        {
            get;
            set;
        }
        public int? FtLo
        {
            get;
            set;
        }
        public int? FtHi
        {
            get;
            set;
        }
        public int? FqT
        {
            get;
            set;
        }
        public int? FqP
        {
            get;
            set;
        }
        public uint? CntIV
        {
            get;
            set;
        }
        public uint? CntKp
        {
            get;
            set;
        }
        public int? Stat
        {
            get;
            set;
        }
        public int Cnting
        {
            get;
            set;
        }
        public int ClrCnt
        {
            get;
            set;
        }
        public int OvLch
        {
            get;
            set;
        }
        #endregion
        //DO
        #region -- DO --
        public int? FSV
        {
            get;
            set;
        }
        public int? PsLo
        {
            get;
            set;
        }
        public int? PsHi
        {
            get;
            set;
        }
        public int? HDT
        {
            get;
            set;
        }
        public int? LDT
        {
            get;
            set;
        }
        public int? ACh
        {
            get;
            set;
        }
        public int? AMd
        {
            get;
            set;
        }
        public int PsCtn
        {
            get;
            set;
        }
        public int PsStop
        {
            get;
            set;
        }
        public int PsIV
        {
            get;
            set;
        }
        #endregion


    }

}
