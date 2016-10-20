using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Text;
using Model;
//20150629
using System.Net;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Threading;
using System.Web.Script.Serialization;


namespace Service
{
    //----------------------------------------------------------------------------------------------//
    // 20151013 Follow AdamApax.NET Class Library VS2008 V8.02.0009 define
    //
    //----------------------------------------------------------------------------------------------//
    public class WISE_AI_RangeInformation
    {
        public static string[] EventStatusEmun = { 
                                               "Fail to provide AI value (UART timeout)",
                                               "Over Range",
                                               "Under Range",
                                               "Open Circuit (Burnout)",
                                               "Reserved",
                                               "Reserved",
                                               "Reserved",
                                               "ADC initializing/Error",
                                               "Reserved",
                                               "Zero/Span Calibration Error"};
        public static string GetUnit(int code)
        {
            switch (code)
            {
                case 259:
                case 260:
                case 261:
                case 262:
                    return "mV";
                case 328:
                case 320:
                case 321:
                case 322:
                case 323:
                case 327:
                case 325:
                    return "V";
                case 384:
                case 385:
                case 386:
                    return "mA";
                case 480: //UDI Mode
                    return "(UDI Mode)";
            }
            return "Invalid Code";
        }
        public static string GetEvent(int evtMask)
        {
            var eventStatus = "";
            var evtEmunLength = EventStatusEmun.Length;
            uint mask = (uint)evtMask;
            for (int i = 0; i <= evtEmunLength; ++i)
            {
                if ((mask & 0x01) == 1)
                {
                    eventStatus += (EventStatusEmun[i] + " ");
                }
                mask >>= 1;
            }
            return eventStatus;
        }
    }
    /// <summary>
    /// module object
    /// </summary>
    public class DeviceData
    {
        public int SL { get; set; }
        public string Id { get; set; }
        public int DIn { get; set; }
        public int DOn { get; set; }
        public int RLAn { get; set; }
        public int AIn { get; set; }
        public int UIn { get; set; }
        public int AOn { get; set; }
        public int Cntn { get; set; }
        public string FwVer { get; set; }
        public string BVer { get; set; }
        //public string ADVer { get; set; }
    }
    public class GetDeviceData
    {
        public DeviceData[] Dev { get; set; }
    }
    //Modbus Info
    public class Modbus_coilconfig
    {
        public int Md { get; set; }
        public int DI { get; set; }
        public int CtS { get; set; }
        public int CtClr { get; set; }
        public int CtOv { get; set; }
        public int Lch { get; set; }
        public int DO { get; set; }
        public int AIHR { get; set; }
        public int AILR { get; set; }
        public int AIB { get; set; }
        public int HAlm { get; set; }
        public int LAlm { get; set; }
        public int GCtClr { get; set; }
    }
    public class Modbus_coillen
    {
        public int DI { get; set; }
        public int CtS { get; set; }
        public int CtClr { get; set; }
        public int CtOv { get; set; }
        public int Lch { get; set; }
        public int DO { get; set; }
        public int AIHR { get; set; }
        public int AILR { get; set; }
        public int AIB { get; set; }
        public int HAlm { get; set; }
        public int LAlm { get; set; }
        public int GCtClr { get; set; }
    }
    public class Modbus_regconfig
    {
        public int Md { get; set; }
        public int DI { get; set; }
        public int CtFq { get; set; }
        public int DO { get; set; }
        public int PsLo { get; set; }
        public int PsHi { get; set; }
        public int PsAV { get; set; }
        public int PsIV { get; set; }
        public int AI { get; set; }
        public int HisH { get; set; }
        public int HisL { get; set; }
        public int AIF { get; set; }
        public int HisHF { get; set; }
        public int HisLF { get; set; }
        public int AIFl { get; set; }
        public int AICd { get; set; }
        public int AlCh { get; set; }
        public int AlSc { get; set; }
        public int AO { get; set; }
        public int AOFl { get; set; }
        public int AOCd { get; set; }
        public int Slew { get; set; }
        public int AOSu { get; set; }
        public int AOSe { get; set; }
        public int AODi { get; set; }
        public int GCLCt { get; set; }
        public int GCLFl { get; set; }
        public int MNm { get; set; }
    }
    public class Modbus_reglen
    {
        public int DI { get; set; }
        public int CtFq { get; set; }
        public int DO { get; set; }
        public int PsLo { get; set; }
        public int PsHi { get; set; }
        public int PsAV { get; set; }
        public int PsIV { get; set; }
        public int AI { get; set; }
        public int HisH { get; set; }
        public int HisL { get; set; }
        public int AIF { get; set; }
        public int HisHF { get; set; }
        public int HisLF { get; set; }
        public int AIFl { get; set; }
        public int AICd { get; set; }
        public int AlCh { get; set; }
        public int AlSc { get; set; }
        public int AO { get; set; }
        public int AOFl { get; set; }
        public int AOCd { get; set; }
        public int Slew { get; set; }
        public int AOSu { get; set; }
        public int AOSe { get; set; }
        public int AODi { get; set; }
        public int GCLCt { get; set; }
        public int GCLFl { get; set; }
        public int MNm { get; set; }
    }









    /// <summary>
    /// Do value object
    /// </summary>
    public class DOValueData
    {
        public int Ch { get; set; }//Channel Number
        public int Md { get; set; }//Mode
        public int Stat { get; set; }//Signal Logic Status
        public uint Val { get; set; }//Channel Value
        public int PsCtn { get; set; }//Pulse Output Continue State
        public int PsStop { get; set; }//Stop Pulse Output
        public int PsIV { get; set; }//Incremental Pulse Output Value

    }
    /// <summary>
    /// Do value object
    /// </summary>
    public class DOSetValueData
    {
        public int Val { get; set; }//Channel Value
    }
    public class DOSlotValueData
    {
        public DOValueData[] DOVal { get; set; }//Array of Digital output values
    }
    /// <summary>
    /// DI value object
    /// </summary>
    public class DIValueData
    {
        public int Ch { get; set; }//Channel Number
        public int Md { get; set; }//Mode
        public uint Val { get; set; }//Channel Value
        public int Stat { get; set; }//Signal Logic Status
        public int Cnting { get; set; }//Start Counter
        public int ClrCnt { get; set; }//Clear the counter value
        public int OvLch { get; set; }//Counter Overflow or Latch Status

    }
    public class DISlotValueData
    {
        public DIValueData[] DIVal { get; set; }//Array of Digital output values
    }

    /// <summary>
    /// AI slot value object
    /// </summary>
    public class AISlotValueData
    {
        public AIValueData[] AIVal { get; set; }//Array of Digital input values
    }
    /// <summary>
    /// AI channel value object
    /// </summary>
    public class AIValueData
    {
        public int Ch { get; set; }//Channel Number
        public int En { get; set; }//Is Enable
        public int Rng { get; set; }//Range
        public uint Val { get; set; }//Channel Raw Value
        public int Eg { get; set; }//Engineering Value
        public int Evt { get; set; }//Event Status
        public int LoA { get; set; }//Low Alarm Status
        public int HiA { get; set; }//High Alarm Status
        public int HVal { get; set; }//Maximum AI Raw Value
        public int HEg { get; set; }//Maximum AI Engineering data
        public int LVal { get; set; }//Minimum AI Raw Value
        public int LEg { get; set; }//Minimum AI Engineering data
        public int SVal { get; set; }//Channel Raw Value After Scaling
        public int ClrH { get; set; }//Clear Maximum AI Value
        public int ClrL { get; set; }//Clear Minimum AI Value
        //20160302 new add
        public float EgF { get; set; }//Channel Engineering data(floating type)
        public float HEgF { get; set; }//Maximum AI Engineering data(floating type)
        public float LEgF { get; set; }//Minimum AI Engineering data(floating type)
    }
    /// <summary>
    /// AI slot value object
    /// </summary>
    public class AISlotConfigData
    {
        public AIConfigData[] AICfg { get; set; }
    }
    /// <summary>
    /// AI channel value object
    /// </summary>
    public class AIConfigData
    {
        public int Ch { get; set; }//Channel Number
        public int En { get; set; }//Is Enable
        public int Rng { get; set; }//Range
        public int EnLA { get; set; }//
        public int EnHA { get; set; }//
        public int LAMd { get; set; }//
        public int HAMd { get; set; }//
        public string LoA { get; set; }//
        public string HiA { get; set; }//
        public string LoS { get; set; }//
        public string HiS { get; set; }//
        public string Tag { get; set; }//
    }
    /// <summary>
    /// AI channel value object
    /// </summary>
    public class AIGenConfigData
    {
        public int Res { get; set; }//
        public int EnB { get; set; }//
        public int BMd { get; set; }//
        public int AiT { get; set; }//
        public int Smp { get; set; }//
        public int AvgM { get; set; }//
    }
    /// <summary>
    /// DO slot value object
    /// </summary>
    public class DOSlotConfigData
    {
        public DOConfigData[] DOCfg { get; set; }
    }
    /// <summary>
    /// AI channel value object
    /// </summary>
    public class DOConfigData
    {
        public int Ch { get; set; }//Channel Number
        public int Md { get; set; }//
        public int FSV { get; set; }//
        public int PsLo { get; set; }//
        public int PsHi { get; set; }//
        public int HDT { get; set; }//
        public int LDT { get; set; }//
        public string Tag { get; set; }//
        public int ACh { get; set; }//
        public int AMd { get; set; }//        
    }
    /// <summary>
    /// DI slot value object
    /// </summary>
    public class DISlotConfigData
    {
        public DIConfigData[] DICfg { get; set; }
    }
    /// <summary>
    /// AI channel value object
    /// </summary>
    public class DIConfigData
    {
        public int Ch { get; set; }//Channel Number
        public int Md { get; set; }//
        public int Inv { get; set; }//
        public int Fltr { get; set; }//
        public int FtLo { get; set; }//
        public int FtHi { get; set; }//
        public int FqT { get; set; }//
        public int FqP { get; set; }//
        public uint CntIV { get; set; }//
        public uint CntKp { get; set; }//
        public string Tag { get; set; }//
    }
    /// <summary>
    /// AO channel value object
    /// </summary>
    public class AOConfigData
    {
        public int Ch { get; set; }//Channel Number
        public int Val { get; set; }//Channel Raw Value
        public int En { get; set; }//Is Enable
        public int Rng { get; set; }//Range
        public string Tag { get; set; }//
    }
    /// <summary>
    /// UI slot value object
    /// </summary>
    public class UISlotConfigData
    {
        public UIConfigData[] UCfg { get; set; }
    }
    /// <summary>
    /// UIO channel value object
    /// </summary>
    public class UIConfigData
    {
        public int Ch { get; set; }//Channel Number
        public int En { get; set; }//Is Enable
        public int Md { get; set; }//Range
    }
    //-----------------------------------------------------------------------------------//
    public delegate void ResquestRespondedCallback(string raw_data);//Define callback function for request has been responded.
    public delegate void ResquestOccurredErrorCallback(Exception e);//Define callback function for request has occurred error.

    
}
