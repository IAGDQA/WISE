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
    /// <summary>
    /// 20150810    Finished subroutine
    /// 20150824    Add to get Modbus config
    /// 20151014    Add new UIO function
    /// 20160315    Add new function for FW update (POST file)
    /// </summary>
    public class HttpReqService //: IHttpReqService
    {
        DeviceModel Device;//20150626 建立一個DeviceModel給所有Service

        private System.Collections.ArrayList BaseListOfIOItems;//20150630 建立一個IO Item List

        private AdvantechHttpWebUtility m_HttpRequest;
        private AdvantechHttpWebUtility m_HttpRequestIO;

        ServiceAction servAct = new ServiceAction();//new 20150824

        bool DevConnFail = false, RefreshMBList = false;

        //Messenger Msg;

        private int m_iUDiRangeCode = 480;//20151013 add

        public HttpReqService()
        {
            //Get IO value
            m_HttpRequest = new AdvantechHttpWebUtility();
            m_HttpRequest.ResquestOccurredError += this.OnGetHttpRequestError;
            m_HttpRequest.ResquestResponded += this.OnGetData;
        }

        //public HttpReqService(Messenger msg)
        //{
        //    Msg = msg;
        //    //Get IO value
        //    m_HttpRequest = new AdvantechHttpWebUtility();
        //    m_HttpRequest.ResquestOccurredError += this.OnGetHttpRequestError;
        //    m_HttpRequest.ResquestResponded += this.OnGetData;
        //    //Get IO Config data
        //    m_HttpRequestIO = new AdvantechHttpWebUtility();
        //    m_HttpRequestIO.ResquestOccurredError += this.OnGetIOConfigHttpRequestError;
        //    m_HttpRequestIO.ResquestResponded += this.OnGetIOConfigData;
        //}

        private string GetURL(string ip, int port, string requestUri)
        {
            return "http://" + ip + ":" + port.ToString() + "/" + requestUri;
        }

        private void OnGetHttpRequestError(Exception e)
        {
            DevConnFail = true;
            //SetTextToTextBox(textBox1, e.Message);
        }

        private void OnGetIOConfigHttpRequestError(Exception e)
        {
            //SetTextToTextBox(textBox1, e.Message);
        }


        #region -- Implement Interface --
        //20150629
        public DeviceModel GetDevice()//DeviceViewModel實作介面
        {
            return Device;
        }

        public bool HttpReqTCP_Connet(string addr)
        {
            return HttpReq_Connet(addr);
        }

        public bool HttpReqTCP_Connet(DeviceModel obj)
        {   //20150629
            bool result = false;

            if (obj.IPAddress == "" || obj.IPAddress == null) return false;

            result = HttpReq_Connet(obj.IPAddress);
            bool flg = false;
            if (result)
            {
                while (!RefreshMBList)
                {
                    if (!flg)//once cycle
                    {
                        GetModbusConfigRequest(ServiceAction.GetModbusCoilsAddr);
                        flg = true;
                    }
                }
                //
                this.InvokeReadStatus();
            }
            return result;
        }

        //Initial creat temp IO list
        public List<IOListData> GetListOfIOItems(string stateFilter)
        {
            List<IOListData> list = new List<IOListData>();
            foreach (var item in BaseListOfIOItems)
            {
                IOModel vd = (IOModel)item;
                if (string.IsNullOrEmpty(stateFilter) /*|| _io.State.ToUpper() == stateFilter.ToUpper()*/)
                {   // need to create item list
                    list.Add(new IOListData()
                    {
                        //UpdateAIConfigUIStatus() <----同步更新
                        //UpdateDIConfigUIStatus() <----同步更新
                        //UpdateDOConfigUIStatus() <----同步更新
                        //Basic
                        Id = vd.Id,
                        Ch = vd.Ch,
                        Tag = vd.Tag,
                        Val = vd.Val,
                        //AI
                        En = vd.En,
                        Rng = vd.Rng,
                        Eg = vd.Eg,
                        EgF = vd.EgF,
                        Evt = vd.Evt,
                        HiA = vd.HiA,
                        LoA = vd.LoA,
                        Val_Eg = vd.Val_Eg,
                        cEn = vd.cEn,
                        cRng = vd.cRng,
                        EnLA = vd.EnLA,
                        EnHA = vd.EnHA,
                        LAMd = vd.LAMd,
                        HAMd = vd.HAMd,
                        cLoA = vd.cLoA,
                        cHiA = vd.cHiA,
                        LoS = vd.LoS,
                        HiS = vd.HiS,
                        // add basic
                        Res = vd.Res,
                        EnB = vd.EnB,
                        BMd = vd.BMd,
                        AiT = vd.AiT,
                        Smp = vd.Smp,
                        AvgM = vd.AvgM,
                        //DI
                        Md = vd.Md,
                        Inv = vd.Inv,
                        Fltr = vd.Fltr,
                        FtLo = vd.FtLo,
                        FtHi = vd.FtHi,
                        FqT = vd.FqT,
                        FqP = vd.FqP,
                        CntIV = vd.CntIV,
                        CntKp = vd.CntKp,
                        Stat = vd.Stat,
                        Cnting = vd.Cnting,
                        ClrCnt = vd.ClrCnt,
                        OvLch = vd.OvLch,
                        //DO
                        FSV = vd.FSV,
                        PsLo = vd.PsLo,
                        PsHi = vd.PsHi,
                        HDT = vd.HDT,
                        LDT = vd.LDT,
                        ACh = vd.ACh,
                        AMd = vd.AMd,
                    });
                }
            }
            return list;
        }
        //20150707 add for AI linear auto test
        public string GetDevRng(string _mdname)
        {
            List<ValueRange> list = new List<ValueRange>();
            if (_mdname == "WISE-4010/LAN")
            { list.Add(ValueRange.mA_4To20); list.Add(ValueRange.mA_0To20); }
            else if (_mdname == "WISE-4012")
            {
                list.Add(ValueRange.mA_4To20); list.Add(ValueRange.mA_0To20);
                list.Add(ValueRange.mA_Neg20To20);
                list.Add(ValueRange.V_Neg10To10); list.Add(ValueRange.V_Neg5To5);
                list.Add(ValueRange.V_Neg1To1); list.Add(ValueRange.mV_Neg500To500);
                list.Add(ValueRange.mV_Neg150To150);
                list.Add(ValueRange.V_0To10); list.Add(ValueRange.V_0To5);
                list.Add(ValueRange.V_0To1); list.Add(ValueRange.mV_0To500);
                list.Add(ValueRange.mV_0To150);
            }
            else return "Failed";

            ValueRange[] listAr = list.ToArray();
            string reStr = "";
            foreach (var _ls in listAr)
            {
                reStr += _ls.ToString() + "\n";
            }
            return reStr;
        }

        public bool UpdateFW(string path)
        {
            bool res = false;
            servAct = ServiceAction.UpdateFW;
            m_HttpRequest.SendPOST_File_Request(Device.Account, Device.Password,
                                            "http://" + Device.IPAddress + "/file_xfer", path);

            return res;
        }

        #region -- IO Service --

        //-- IO List Service --
        public IOModel GetIO(int _id)
        {
            IOModel _RtnInfo = new IOModel();
            foreach (var item in BaseListOfIOItems)
            {
                var readInfo = (IOModel)item;
                if (readInfo.Id.Equals(_id))
                {
                    _RtnInfo = readInfo;
                    break;
                }
            }
            return _RtnInfo;
        }

        public IOModel GetIOConfig(int _id)
        {
            IOModel _RtnInfo = new IOModel();
            //實作IO config api
            GetIOConfigRequest(_id);
            //Thread.Sleep(100);
            foreach (var item in BaseListOfIOItems)
            {
                var readInfo = (IOModel)item;
                if (readInfo.Id.Equals(_id))
                {
                    _RtnInfo = readInfo;
                    break;
                }
            }
            return _RtnInfo;
        }

        public void UpdateIOConfig(IOModel data)
        {
            object obj = new object();
            if (data.Id >= (int)DevIDDefine.AI)
            {
                obj = new AIConfigData()
                {
                    Ch = data.Ch,
                    Tag = data.Tag,
                    En = data.cEn,
                    Rng = data.cRng,
                    EnLA = data.EnLA.Value,
                    EnHA = data.EnHA.Value,
                    LAMd = data.LAMd.Value,
                    HAMd = data.HAMd.Value,
                    LoA = data.cLoA,
                    HiA = data.cHiA,
                    LoS = data.LoS,
                    HiS = data.HiS,
                };
            }
            else if (data.Id >= (int)DevIDDefine.DO)
            {
                obj = new DOConfigData()
                {
                    Ch = data.Ch,
                    Tag = data.Tag,
                    Md = data.Md.Value,
                    FSV = data.FSV.Value,
                    PsLo = data.PsLo.Value,
                    PsHi = data.PsHi.Value,
                    HDT = data.HDT.Value,
                    LDT = data.LDT.Value,
                    ACh = data.ACh.Value,
                    AMd = data.AMd.Value,
                };
            }
            else
            {
                obj = new DIConfigData()
                {
                    Ch = data.Ch,
                    Tag = data.Tag,
                    Md = data.Md.Value,
                    Inv = data.Inv.Value,
                    Fltr = data.Fltr.Value,
                    FtLo = data.FtLo.Value,
                    FtHi = data.FtHi.Value,
                    FqT = data.FqT.Value,
                    FqP = data.FqP.Value,
                    CntIV = data.CntIV.Value,
                    CntKp = data.CntKp.Value,
                };
            }
            SetIOConfigRequest(data.Id.Value, data.Ch, obj);
        }
        //20151014 new add
        //public void UpdateUIOConfig(IOModel data)
        //{
        //    object obj = new object();
        //    obj = new UIOConfigData()
        //    {
        //        Ch = data.Ch,
        //        Md = data.Md.Value,
        //        En = data.cEn,
        //    };
        //}

        public void UpdateIOValue(IOModel data)
        {
            object obj = new object();
            if (data.Id >= (int)DevIDDefine.AI)
            {
                obj = new AIValueData()
                {
                    Ch = data.Ch,
                    LoA = data.LoA,
                    HiA = data.HiA,
                };
            }
            else if (data.Id >= (int)DevIDDefine.DO)
            {
                obj = new DOValueData()
                {
                    Ch = data.Ch,
                    Val = data.Val,
                    PsCtn = data.PsCtn,
                    PsStop = data.PsStop,
                    PsIV = data.PsIV,
                };
            }
            else
            {
                obj = new DIValueData()
                {
                    Ch = data.Ch,
                    Val = data.Val,
                    Cnting = data.Cnting,
                    ClrCnt = data.ClrCnt,
                    OvLch = data.OvLch,
                };
            }

            SetIOValueRequest(data.Id.Value, data.Ch, obj);
        }


        #endregion //IO Service

        #endregion


        /// <summary>
        /// ////////////////////////////////////////////////////////////////////
        /// </summary>
        private bool HttpReq_Connet(string addr)
        {
            bool res = false;
            Device = new DeviceModel()//20150626 建立一個DeviceModel給所有Service
            {
                IPAddress = addr,
                Account = "root",
                Password = "00000000",
                Port = 80,
                SlotNum = 0,
                ModbusAddr = 1,
                ModbusTimeOut = 3000,
            };

            GetDevInfoRequest();//實作HttpReq connect
            RefreshMBList = false;
            //response to Device.IPAddress event
            int time_out = 0;
            while (!Device.ConnFlg)//prepose is delay to wait response event.
            {
                if (DevConnFail)
                { DevConnFail = false; res = false; break; }
                else res = true;
                servAct = ServiceAction.Connection;
                System.Threading.Thread.Sleep(1000);
                time_out++;
                if (time_out > 4) { res = false; break; }
            }

            return res;
        }

        private void OnGetData(string rawData)//Feedback Http request
        {
            if (!Device.ConnFlg) servAct = ServiceAction.Connection;//Priority High

            switch (servAct)
            {
                case ServiceAction.Connection:
                    var Obj01 = AdvantechHttpWebUtility.ParserJsonToObj<GetDeviceData>(rawData);
                    UpdateDevUIStatus(Obj01);
                    break;
                case ServiceAction.GetModbusCoilsAddr:
                    var Obj02 = AdvantechHttpWebUtility.ParserJsonToObj<Modbus_coilconfig>(rawData);
                    UpdateMbCoilAddr(Obj02);
                    break;
                case ServiceAction.GetModbusCoilsLength:
                    var Obj03 = AdvantechHttpWebUtility.ParserJsonToObj<Modbus_coillen>(rawData);
                    UpdateMbCoilLen(Obj03);
                    break;
                case ServiceAction.GetModbusRegsAddr:
                    var Obj04 = AdvantechHttpWebUtility.ParserJsonToObj<Modbus_regconfig>(rawData);
                    UpdateMbRegAddr(Obj04);
                    break;
                case ServiceAction.GetModbusRegsLength:
                    var Obj05 = AdvantechHttpWebUtility.ParserJsonToObj<Modbus_reglen>(rawData);
                    UpdateMbRegLen(Obj05);
                    break;
                //                
                case ServiceAction.GetAIVal:
                    var Obj06 = AdvantechHttpWebUtility.ParserJsonToObj<AISlotValueData>(rawData);
                    UpdateAIUIStatus(Obj06);
                    if (Device.DiTotal > 0)
                        GetDIValue();
                    else if (Device.DoTotal > 0 || Device.RLaTotal > 0)
                        GetDOValue();
                    else
                        InvokeReadStatus();
                    break;
                case ServiceAction.GetDOVal:
                    var Obj07 = AdvantechHttpWebUtility.ParserJsonToObj<DOSlotValueData>(rawData);
                    UpdateDOUIStatus(Obj07);
                    InvokeReadStatus();
                    break;
                case ServiceAction.GetDIVal:
                    var Obj08 = AdvantechHttpWebUtility.ParserJsonToObj<DISlotValueData>(rawData);
                    UpdateDIUIStatus(Obj08);
                    if (Device.DoTotal > 0 || Device.RLaTotal > 0)//has DO
                        GetDOValue();
                    else
                        InvokeReadStatus();
                    break;
                //
                case ServiceAction.GetUInMod:
                    var Obj09 = AdvantechHttpWebUtility.ParserJsonToObj<UISlotConfigData>(rawData);
                    UpdateUIOUIStatus(Obj09);
                    RefreshIOStatus();
                    break;
                //Jammy Add 20160315
                case ServiceAction.UpdateFW:
                    
                    break;
                default:

                    break;
            }
            //finished
        }

        private void GetMBCoilsAddrPrc()
        {
            servAct = ServiceAction.GetModbusCoilsAddr;
            GetModbusConfigRequest(servAct);
        }
        private void GetMBCoilsLenPrc()
        {
            servAct = ServiceAction.GetModbusCoilsLength;
            GetModbusConfigRequest(servAct);
        }
        private void GetMBRegsAddrPrc()
        {
            servAct = ServiceAction.GetModbusRegsAddr;
            GetModbusConfigRequest(servAct);
        }
        private void GetMBRegsLenPrc()
        {
            servAct = ServiceAction.GetModbusRegsLength;
            GetModbusConfigRequest(servAct);
        }

        private void InvokeReadStatus()//start to read IO invoke
        {
            int m_iPollingTime = 100;
            Thread.Sleep(m_iPollingTime);//m_iPollingTime
            //
            //NotificationResult result = Msg.NotifyColleagues(MessageTypes.MSG_IOITEMS_UPDATE_REQ, BaseListOfIOItems);
            //if (result == NotificationResult.MessageNotRegistered || result == NotificationResult.MessageRegisteredNotHandled)
            //    Msg.NotifyColleagues(MessageTypes.MSG_IOITEMS_UPDATE_REQ, BaseListOfIOItems);

            RefreshIOStatus();
        }

        private void RefreshIOStatus()
        {
            //if (!OneFlg)
            //    GetModbusConfigRequest(ServiceAction.GetModbusCoilsAddr);     
            if (Device.UInTotal > 0)
            {
                if (UIOMod == 99) GetMod();
                else if (UIOMod == 0)
                    GetAIValue();
                else if (UIOMod == 1)
                    GetDIValue();
                //else
                //    GetDOValue();
            }
            else
            {
                if (Device.AiTotal > 0)
                    GetAIValue();
                else if (Device.DiTotal > 0)
                    GetDIValue();
                else
                    GetDOValue();
            }
        }

        //Request Cmd
        private void GetDevInfoRequest()
        {
            servAct = ServiceAction.Connection;
            m_HttpRequest.SendGETRequest(Device.Account, Device.Password,
                                            "http://" + Device.IPAddress + "/profile");
        }
        private void GetModbusConfigRequest(ServiceAction req)
        {
            if (req == ServiceAction.GetModbusCoilsAddr)
            {
                servAct = ServiceAction.GetModbusCoilsAddr;
                m_HttpRequest.SendGETRequest(Device.Account, Device.Password,
                                                "http://" + Device.IPAddress + "/modbus_coilconfig");
            }
            else if (req == ServiceAction.GetModbusCoilsLength)
            {
                servAct = ServiceAction.GetModbusCoilsLength;
                m_HttpRequest.SendGETRequest(Device.Account, Device.Password,
                                                "http://" + Device.IPAddress + "/modbus_coillen");
            }
            else if (req == ServiceAction.GetModbusRegsAddr)
            {
                servAct = ServiceAction.GetModbusRegsAddr;
                m_HttpRequest.SendGETRequest(Device.Account, Device.Password,
                                                "http://" + Device.IPAddress + "/modbus_regconfig");
            }
            else if (req == ServiceAction.GetModbusRegsLength)
            {
                servAct = ServiceAction.GetModbusRegsLength;
                m_HttpRequest.SendGETRequest(Device.Account, Device.Password,
                                                "http://" + Device.IPAddress + "/modbus_reglen");
            }
        }
        int get_IO_request_id = 0;
        private void GetIOConfigRequest(int _id)
        {   //get Multi Channel Request
            get_IO_request_id = _id;
            if (_id >= (int)DevIDDefine.AI)
                m_HttpRequestIO.SendGETRequest(Device.Account, Device.Password, GetURL(Device.IPAddress, Device.Port
                               , WISE_RESTFUL_URI.ai_config.ToString() + "/slot_" + Device.SlotNum));
            else if (_id >= (int)DevIDDefine.DO)
                m_HttpRequestIO.SendGETRequest(Device.Account, Device.Password, GetURL(Device.IPAddress, Device.Port
                               , WISE_RESTFUL_URI.do_config.ToString() + "/slot_" + Device.SlotNum));
            else
                m_HttpRequestIO.SendGETRequest(Device.Account, Device.Password, GetURL(Device.IPAddress, Device.Port
                               , WISE_RESTFUL_URI.di_config.ToString() + "/slot_" + Device.SlotNum));
        }

        private void SetIOConfigRequest(int _id, int _ch, object _obj)
        {
            JavaScriptSerializer serializer = new JavaScriptSerializer();
            //AdvantechHttpWebUtility httpRequest = new AdvantechHttpWebUtility();
            string sz_Jsonify = serializer.Serialize(_obj);
            try
            {
                if (_id >= (int)DevIDDefine.AI)//AI
                    m_HttpRequestIO.SendPATCHRequest(Device.Account, Device.Password, GetURL(Device.IPAddress, Device.Port
                                    , WISE_RESTFUL_URI.ai_config + "/slot_" + Device.SlotNum
                                    + "/ch_" + _ch), sz_Jsonify);
                else if (_id >= (int)DevIDDefine.DO)//DO
                    m_HttpRequestIO.SendPATCHRequest(Device.Account, Device.Password, GetURL(Device.IPAddress, Device.Port
                                    , WISE_RESTFUL_URI.do_config + "/slot_" + Device.SlotNum
                                    + "/ch_" + _ch), sz_Jsonify);
                else//DI
                    m_HttpRequestIO.SendPATCHRequest(Device.Account, Device.Password, GetURL(Device.IPAddress, Device.Port
                                    , WISE_RESTFUL_URI.di_config + "/slot_" + Device.SlotNum
                                    + "/ch_" + _ch), sz_Jsonify);
                //
                GetIOConfigRequest(_id);
            }
            catch (Exception e)
            {
                //OnGetDOHttpRequestError(e);
            }
            finally
            {
                System.GC.Collect();
            }
        }

        private void SetIOValueRequest(int _id, int _ch, object _obj)
        {
            JavaScriptSerializer serializer = new JavaScriptSerializer();
            AdvantechHttpWebUtility httpRequest = new AdvantechHttpWebUtility();
            string sz_Jsonify = serializer.Serialize(_obj);
            try
            {
                if (_id >= (int)DevIDDefine.AI)//AI
                    httpRequest.SendPATCHRequest(Device.Account, Device.Password, GetURL(Device.IPAddress, Device.Port
                                    , WISE_RESTFUL_URI.ai_value + "/slot_" + Device.SlotNum
                                    + "/ch_" + _ch), sz_Jsonify);
                else if (_id >= (int)DevIDDefine.DO)//DO
                    httpRequest.SendPATCHRequest(Device.Account, Device.Password, GetURL(Device.IPAddress, Device.Port
                                    , WISE_RESTFUL_URI.do_value + "/slot_" + Device.SlotNum
                                    + "/ch_" + _ch), sz_Jsonify);
                else//DI
                    httpRequest.SendPATCHRequest(Device.Account, Device.Password, GetURL(Device.IPAddress, Device.Port
                                    , WISE_RESTFUL_URI.di_value + "/slot_" + Device.SlotNum
                                    + "/ch_" + _ch), sz_Jsonify);
            }
            catch (Exception e)
            {
                //OnGetDOHttpRequestError(e);
            }
            finally
            {
                System.GC.Collect();
            }
        }
        //------------------------------------------------------------------------//
        #region ---- Get Value ----
        private void GetDIValue()
        {
            servAct = ServiceAction.GetDIVal;
            m_HttpRequest.SendGETRequest(Device.Account, Device.Password, GetURL(Device.IPAddress, Device.Port
                               , WISE_RESTFUL_URI.di_value.ToString() + "/slot_" + Device.SlotNum));
        }
        private void GetDOValue()
        {
            servAct = ServiceAction.GetDOVal;
            m_HttpRequest.SendGETRequest(Device.Account, Device.Password, GetURL(Device.IPAddress, Device.Port
                               , WISE_RESTFUL_URI.do_value.ToString() + "/slot_" + Device.SlotNum));
        }
        private void GetAIValue()
        {
            servAct = ServiceAction.GetAIVal;
            m_HttpRequest.SendGETRequest(Device.Account, Device.Password, GetURL(Device.IPAddress, Device.Port
                               , WISE_RESTFUL_URI.ai_value.ToString() + "/slot_" + Device.SlotNum));
        }
        //20151014 new add
        private void GetMod()
        {
            servAct = ServiceAction.GetUInMod;
            m_HttpRequest.SendGETRequest(Device.Account, Device.Password, GetURL(Device.IPAddress, Device.Port
                               , WISE_RESTFUL_URI.uio_config.ToString() + "/slot_" + Device.SlotNum));
        }
        #endregion
        //------------------------------------------------------------------------//
        private void OnGetIOConfigData(string rawData)
        {
            if (get_IO_request_id >= (int)DevIDDefine.AI)
            {
                var dateObj = AdvantechHttpWebUtility.ParserJsonToObj<AISlotConfigData>(rawData);
                UpdateAIConfigUIStatus(dateObj);
            }
            else if (get_IO_request_id >= (int)DevIDDefine.DO)
            {
                var dateObj = AdvantechHttpWebUtility.ParserJsonToObj<DOSlotConfigData>(rawData);
                UpdateDOConfigUIStatus(dateObj);
            }
            else
            {
                var dateObj = AdvantechHttpWebUtility.ParserJsonToObj<DISlotConfigData>(rawData);
                UpdateDIConfigUIStatus(dateObj);
            }
        }
        //------------------------------------------------------------------------//
        #region ---- Update UI ----
        private void UpdateDevUIStatus(GetDeviceData devdata)
        {
            BaseListOfIOItems = new System.Collections.ArrayList();
            try
            {
                if (devdata.Dev != null)
                {
                    for (int i = 0; i < devdata.Dev.Length; ++i)
                    {   //將資料存入Device
                        Device.ModuleType = devdata.Dev[i].Id.ToString();
                        Device.FirmwareVer = devdata.Dev[i].FwVer.ToString();
                        //Device.Account = devdata.Dev[i]..ToString();
                        //Device.Password = devdata.Dev[i].FwVer.ToString();
                        Device.DiTotal = devdata.Dev[i].DIn;
                        Device.DoTotal = devdata.Dev[i].DOn;
                        Device.RLaTotal = devdata.Dev[i].RLAn;
                        Device.AiTotal = devdata.Dev[i].AIn;
                        Device.UInTotal = devdata.Dev[i].UIn;
                        Device.AoTotal = devdata.Dev[i].AOn;
                        Device.CntTotal = devdata.Dev[i].Cntn;
                        Device.ConnFlg = true;
                    }
                }
                else
                {
                    throw new Exception("Parser Conn Data Fail");
                }
            }
            catch (Exception e)
            {
                //OnGetDevHttpRequestError(e);
            }
            finally
            {
                System.GC.Collect();
            }
        }

        private void UpdateMbCoilAddr(Modbus_coilconfig conf)
        {
            try
            {
                if (Device != null)
                {
                    Device.MbCoils = new ModbusCoilModel()
                    {
                        Md = conf.Md,
                        DI = conf.DI,
                        CtS = conf.CtS,
                        CtClr = conf.CtClr,
                        CtOv = conf.CtOv,
                        Lch = conf.Lch,
                        DO = conf.DO,
                        AIHR = conf.AIHR,
                        AILR = conf.AILR,
                        AIB = conf.AIB,
                        HAlm = conf.HAlm,
                        LAlm = conf.LAlm,
                        GCtClr = conf.GCtClr,
                    };
                }
                else
                {
                    throw new Exception("Parser Conn Data Fail");
                }
            }
            catch (Exception e)
            {
                //OnGetDevHttpRequestError(e);
            }
            finally
            {
                System.GC.Collect();
            }
            //
            //servAct = ServiceAction.GetModbusCoilsLength;
            GetModbusConfigRequest(ServiceAction.GetModbusCoilsLength);
        }
        private void UpdateMbCoilLen(Modbus_coillen conf)
        {
            try
            {
                if (Device != null && Device.MbCoils != null)
                {
                    Device.MbCoils.lenDI = conf.DI;
                    Device.MbCoils.lenCtS = conf.CtS;
                    Device.MbCoils.lenCtClr = conf.CtClr;
                    Device.MbCoils.lenCtOv = conf.CtOv;
                    Device.MbCoils.lenLch = conf.Lch;
                    Device.MbCoils.lenDO = conf.DO;
                    Device.MbCoils.lenAIHR = conf.AIHR;
                    Device.MbCoils.lenAILR = conf.AILR;
                    Device.MbCoils.lenAIB = conf.AIB;
                    Device.MbCoils.lenHAlm = conf.HAlm;
                    Device.MbCoils.lenLAlm = conf.LAlm;
                    Device.MbCoils.lenGCtClr = conf.GCtClr;
                }
                else
                {
                    throw new Exception("Parser Conn Data Fail");
                }
            }
            catch (Exception e)
            {
                //OnGetDevHttpRequestError(e);
            }
            finally
            {
                System.GC.Collect();
            }
            //
            //servAct = ServiceAction.GetModbusRegsAddr;
            GetModbusConfigRequest(ServiceAction.GetModbusRegsAddr);
        }
        private void UpdateMbRegAddr(Modbus_regconfig conf)
        {
            try
            {
                if (Device != null)
                {
                    Device.MbRegs = new ModbusRegModel()
                    {
                        Md = conf.Md,
                        DI = conf.DI,
                        CtFq = conf.CtFq,
                        DO = conf.DO,
                        PsLo = conf.PsLo,
                        PsHi = conf.PsHi,
                        PsAV = conf.PsAV,
                        PsIV = conf.PsIV,
                        AI = conf.AI,
                        HisH = conf.HisH,
                        HisL = conf.HisL,
                        AIF = conf.AIF,
                        HisHF = conf.HisHF,
                        HisLF = conf.HisLF,
                        AIFl = conf.AIFl,
                        AICd = conf.AICd,
                        AlCh = conf.AlCh,
                        AlSc = conf.AlSc,
                        AO = conf.AO,
                        AOFl = conf.AOFl,
                        AOCd = conf.AOCd,
                        Slew = conf.Slew,
                        AOSu = conf.AOSu,
                        AOSe = conf.AOSe,
                        AODi = conf.AODi,
                        GCLCt = conf.GCLCt,
                        GCLFl = conf.GCLFl,
                        MNm = conf.MNm,
                    };
                }
                else
                {
                    throw new Exception("Parser Conn Data Fail");
                }
            }
            catch (Exception e)
            {
                //OnGetDevHttpRequestError(e);
            }
            finally
            {
                System.GC.Collect();
            }
            //
            //servAct = ServiceAction.GetModbusRegsLength;
            GetModbusConfigRequest(ServiceAction.GetModbusRegsLength);
        }
        private void UpdateMbRegLen(Modbus_reglen conf)
        {
            try
            {
                if (Device != null && Device.MbRegs != null)
                {
                    Device.MbRegs.lenDI = conf.DI;
                    Device.MbRegs.lenCtFq = conf.CtFq;
                    Device.MbRegs.lenDO = conf.DO;
                    Device.MbRegs.lenPsLo = conf.PsLo;
                    Device.MbRegs.lenPsHi = conf.PsHi;
                    Device.MbRegs.lenPsAV = conf.PsAV;
                    Device.MbRegs.lenPsIV = conf.PsIV;
                    Device.MbRegs.lenAI = conf.AI;
                    Device.MbRegs.lenHisH = conf.HisH;
                    Device.MbRegs.lenHisL = conf.HisL;
                    Device.MbRegs.lenAIF = conf.AIF;
                    Device.MbRegs.lenHisHF = conf.HisHF;
                    Device.MbRegs.lenHisLF = conf.HisLF;
                    Device.MbRegs.lenAIFl = conf.AIFl;
                    Device.MbRegs.lenAICd = conf.AICd;
                    Device.MbRegs.lenAlCh = conf.AlCh;
                    Device.MbRegs.lenAlSc = conf.AlSc;
                    Device.MbRegs.lenAO = conf.AO;
                    Device.MbRegs.lenAOFl = conf.AOFl;
                    Device.MbRegs.lenAOCd = conf.AOCd;
                    Device.MbRegs.lenSlew = conf.Slew;
                    Device.MbRegs.lenAOSu = conf.AOSu;
                    Device.MbRegs.lenAOSe = conf.AOSe;
                    Device.MbRegs.lenAODi = conf.AODi;
                    Device.MbRegs.lenGCLCt = conf.GCLCt;
                    Device.MbRegs.lenGCLFl = conf.GCLFl;
                    Device.MbRegs.lenMNm = conf.MNm;
                }
                else
                {
                    throw new Exception("Parser Conn Data Fail");
                }
            }
            catch (Exception e)
            {
                //OnGetDevHttpRequestError(e);
            }
            finally
            {
                System.GC.Collect();
            }
            //
            RefreshMBList = true;
        }

        private void UpdateDIUIStatus(DISlotValueData di_values)//Event to update data
        {
            try
            {
                if (di_values.DIVal != null)
                {
                    for (int i = 0; i < di_values.DIVal.Length; ++i)
                    {
                        IOModel BfMdfIOCh = GetIO(i + (int)DevIDDefine.DI);
                        if (BfMdfIOCh.Id != null && BfMdfIOCh.Ch == i)
                        {
                            var temp = BfMdfIOCh;
                            temp.Val = di_values.DIVal[i].Val;
                            temp.Md = di_values.DIVal[i].Md;
                            temp.Stat = di_values.DIVal[i].Stat;
                            temp.Cnting = di_values.DIVal[i].Cnting;
                            temp.OvLch = di_values.DIVal[i].OvLch;

                            int Idx = BaseListOfIOItems.IndexOf(BfMdfIOCh);
                            BaseListOfIOItems.Remove(BfMdfIOCh);
                            BaseListOfIOItems.Insert(Idx, temp);
                        }
                        else if (BfMdfIOCh.Id == null)
                        {
                            IOModel _BfMdfIOCh = new IOModel()
                            {
                                Id = i + (int)DevIDDefine.DI,
                                Ch = i,
                            };
                            BaseListOfIOItems.Add(_BfMdfIOCh);
                            //20150702 one shot to scan all IO config
                            int _id = _BfMdfIOCh.Id.Value;
                            GetIOConfigRequest(_id);
                        }
                    }
                }
                else
                {
                    throw new Exception("Parser DI Data Fail");
                }
            }
            catch (Exception e)
            {
                servAct = ServiceAction.Idel; return;//OnGetDIHttpRequestError(e);
            }
            finally
            {
                System.GC.Collect();
            }
        }

        private void UpdateDOUIStatus(DOSlotValueData do_values)
        {
            try
            {
                if (do_values.DOVal != null)
                {
                    for (int i = 0; i < do_values.DOVal.Length; ++i)
                    {
                        IOModel BfMdfIOCh = GetIO(i + (int)DevIDDefine.DO);
                        if (BfMdfIOCh.Id != null && BfMdfIOCh.Ch == i)
                        {
                            var temp = BfMdfIOCh;
                            temp.Val = do_values.DOVal[i].Val;
                            temp.Stat = do_values.DOVal[i].Stat;
                            temp.PsCtn = do_values.DOVal[i].PsCtn;
                            temp.PsIV = do_values.DOVal[i].PsIV;

                            int Idx = BaseListOfIOItems.IndexOf(BfMdfIOCh);
                            BaseListOfIOItems.Remove(BfMdfIOCh);
                            BaseListOfIOItems.Insert(Idx, temp);
                        }
                        else if (BfMdfIOCh.Id == null)
                        {
                            IOModel _BfMdfIOCh = new IOModel()
                            {
                                Id = (int)DevIDDefine.DO + i,
                                Ch = i,
                            };
                            BaseListOfIOItems.Add(_BfMdfIOCh);
                            //20150702 one shot to scan all IO config
                            int _id = _BfMdfIOCh.Id.Value;
                            GetIOConfigRequest(_id);
                        }
                    }
                }
                else
                {
                    throw new Exception("Parser DO Data Fail");
                }
            }
            catch (Exception e)
            {
                servAct = ServiceAction.Idel; return;//OnGetDOHttpRequestError(e);
            }
            finally
            {
                System.GC.Collect();
            }

        }

        private void UpdateAIUIStatus(AISlotValueData ai_values)
        {
            try
            {
                if (ai_values.AIVal != null)
                {
                    for (int i = 0; i < ai_values.AIVal.Length - 1; ++i)
                    {
                        IOModel BfMdfIOCh = GetIO(i + (int)DevIDDefine.AI);
                        if (BfMdfIOCh.Id != null && BfMdfIOCh.Ch == i)
                        {
                            var temp = BfMdfIOCh;
                            if (ai_values.AIVal[i].En > 0)//removed 20150713
                            {
                                int rangeCode = ai_values.AIVal[i].Rng;
                                int aiEvent = ai_values.AIVal[i].Evt;

                                if (rangeCode == m_iUDiRangeCode)//20151013 new
                                {
                                    temp.Val = ai_values.AIVal[i].Val <= 0 ? 0 : (uint)1;
                                    temp.Val_Eg = (ai_values.AIVal[i].Val <= 0) ? "Low" : "High"
                                                    + " " + WISE_AI_RangeInformation.GetUnit(rangeCode);
                                }
                                else
                                {
                                    temp.Val = ai_values.AIVal[i].Val;
                                    string unit = WISE_AI_RangeInformation.GetUnit(rangeCode);
                                    float engineerFloatValue = ai_values.AIVal[i].EgF;//WISE-4012 vA101B00 之後改為EgF
                                    if (unit == "V" || unit == "A")
                                    {
                                        engineerFloatValue = engineerFloatValue / 1000;
                                        temp.Val_Eg = (ai_values.AIVal[i].Evt > 0) ? WISE_AI_RangeInformation.GetEvent(aiEvent) : ((Math.Truncate(engineerFloatValue * 10000) / 10000).ToString("0.0000") + " " + unit);
                                    }
                                    else//mV or mA
                                        temp.Val_Eg = (ai_values.AIVal[i].Evt > 0) ? WISE_AI_RangeInformation.GetEvent(aiEvent) : (engineerFloatValue + " " + unit);
                                    //WISE-4012 vA101B00 之後改為EgF
                                    //temp.Val_Eg = (ai_values.AIVal[i].Evt > 0) ? WISE_AI_RangeInformation.GetEvent(aiEvent)
                                    //                    : ((ai_values.AIVal[i].EgF).ToString("0.000") + " "
                                    //                    + WISE_AI_RangeInformation.GetUnit(rangeCode));
                                }
                                temp.En = ai_values.AIVal[i].En;
                                temp.Rng = ai_values.AIVal[i].Rng;
                                temp.Eg = ai_values.AIVal[i].Eg;
                                temp.EgF = ai_values.AIVal[i].EgF;
                                temp.Evt = ai_values.AIVal[i].Evt;
                                temp.LoA = ai_values.AIVal[i].LoA;
                                temp.HiA = ai_values.AIVal[i].HiA;

                                //SetTextToTextBox(m_AITextBoxList[textboxIndex], value);
                            }
                            else
                            {
                                temp.Val = 0;
                                temp.Val_Eg = "Disable";
                                temp.En = ai_values.AIVal[i].En;
                                temp.Rng = ai_values.AIVal[i].Rng;
                            }
                            int Idx = BaseListOfIOItems.IndexOf(BfMdfIOCh);
                            BaseListOfIOItems.Remove(BfMdfIOCh);
                            BaseListOfIOItems.Insert(Idx, temp);
                        }
                        else if (BfMdfIOCh.Id == null)
                        {
                            IOModel _BfMdfIOCh = new IOModel()
                            {
                                Id = (int)DevIDDefine.AI + i,
                                Ch = i,
                            };
                            BaseListOfIOItems.Add(_BfMdfIOCh);
                            //20150702 one shot to scan all IO config
                            int _id = _BfMdfIOCh.Id.Value;
                            GetIOConfigRequest(_id);
                        }
                    }
                }
                else
                    throw new Exception("Parser AI Data Fail");
            }
            catch (Exception e)
            {
                servAct = ServiceAction.Idel; return;//OnGetAIHttpRequestError(e);
            }
            finally
            {
                System.GC.Collect();
            }

        }
        int UIOMod = 99;//20151014 new add
        private void UpdateUIOUIStatus(UISlotConfigData uio_config)//20151014 new add
        {
            try
            {
                if (uio_config.UCfg != null)
                {   //只判斷CH0是否為AI或DI
                    if (uio_config.UCfg[0].Md == 2)
                        UIOMod = 2;
                    else if (uio_config.UCfg[0].Md == 1)
                        UIOMod = 1;
                    else
                        UIOMod = 0;
                }
            }
            catch (Exception e)
            {
                servAct = ServiceAction.Idel; return;
            }
            finally
            {
                System.GC.Collect();
            }
        }
        #endregion
        //--------- Config -------------------------------------------------------//
        #region ---- Config ----
        private void UpdateAIGenConfigUIStatus(AIGenConfigData ai_cfgs)
        {
            try
            {
                Device.Res = ai_cfgs.Res;
                Device.EnB = ai_cfgs.EnB;
                Device.BMd = ai_cfgs.BMd;
                Device.AiT = ai_cfgs.AiT;
                Device.Smp = ai_cfgs.Smp;
                Device.AvgM = ai_cfgs.AvgM;
            }
            catch (Exception e)
            {
                OnGetIOConfigHttpRequestError(e);
            }
            finally
            {
                System.GC.Collect();
            }
        }
        private void UpdateAIConfigUIStatus(AISlotConfigData ai_cfgs)
        {
            int i = 0;
            try
            {
                if (ai_cfgs.AICfg != null)
                {
                    while (GetIO(i + (int)DevIDDefine.AI) != null)
                    {
                        var _IOdata = GetIO(i + (int)DevIDDefine.AI);
                        //MainController.HttpReqService.GetIOConfig() <----同步更新
                        _IOdata.cEn = ai_cfgs.AICfg[i].En;
                        _IOdata.Tag = ai_cfgs.AICfg[i].Tag;
                        _IOdata.cRng = ai_cfgs.AICfg[i].Rng;
                        _IOdata.EnLA = ai_cfgs.AICfg[i].EnLA;
                        _IOdata.EnHA = ai_cfgs.AICfg[i].EnHA;
                        _IOdata.LAMd = ai_cfgs.AICfg[i].LAMd;
                        _IOdata.HAMd = ai_cfgs.AICfg[i].HAMd;
                        _IOdata.cLoA = ai_cfgs.AICfg[i].LoA;
                        _IOdata.cHiA = ai_cfgs.AICfg[i].HiA;
                        _IOdata.LoS = ai_cfgs.AICfg[i].LoS;
                        _IOdata.HiS = ai_cfgs.AICfg[i].HiS;
                        //add basic function
                        _IOdata.Res = Device.Res;
                        _IOdata.EnB = Device.EnB;
                        _IOdata.BMd = Device.BMd;
                        _IOdata.AiT = Device.AiT;
                        _IOdata.Smp = Device.Smp;
                        _IOdata.AvgM = Device.AvgM;
                        i++;
                    }
                }
                else
                    throw new Exception("Parser AI Data Fail");
            }
            catch (Exception e)
            {
                OnGetIOConfigHttpRequestError(e);
            }
            finally
            {
                System.GC.Collect();
            }
        }
        private void UpdateDIConfigUIStatus(DISlotConfigData di_cfgs)
        {
            int i = 0;
            try
            {
                if (di_cfgs.DICfg != null)
                {
                    while (GetIO(i + (int)DevIDDefine.DI) != null)
                    {
                        var _IOdata = GetIO(i + (int)DevIDDefine.DI);
                        //MainController.HttpReqService.GetIOConfig() <----同步更新
                        _IOdata.Md = di_cfgs.DICfg[i].Md;
                        _IOdata.Inv = di_cfgs.DICfg[i].Inv;
                        _IOdata.Fltr = di_cfgs.DICfg[i].Fltr;
                        _IOdata.FtLo = di_cfgs.DICfg[i].FtLo;
                        _IOdata.FtHi = di_cfgs.DICfg[i].FtHi;
                        _IOdata.FqT = di_cfgs.DICfg[i].FqT;
                        _IOdata.FqP = di_cfgs.DICfg[i].FqP;
                        _IOdata.CntIV = di_cfgs.DICfg[i].CntIV;
                        _IOdata.CntKp = di_cfgs.DICfg[i].CntKp;
                        _IOdata.Tag = di_cfgs.DICfg[i].Tag;
                        i++;
                    }
                }
                else
                    throw new Exception("Parser AI Data Fail");
            }
            catch (Exception e)
            {
                OnGetIOConfigHttpRequestError(e);
            }
            finally
            {
                System.GC.Collect();
            }
        }
        private void UpdateDOConfigUIStatus(DOSlotConfigData do_cfgs)
        {
            int i = 0;
            try
            {
                if (do_cfgs.DOCfg != null)
                {
                    while (GetIO(i + (int)DevIDDefine.DO) != null)
                    {
                        var _IOdata = GetIO(i + (int)DevIDDefine.DO);
                        //MainController.HttpReqService.GetIOConfig() <----同步更新
                        _IOdata.Md = do_cfgs.DOCfg[i].Md;
                        _IOdata.FSV = do_cfgs.DOCfg[i].FSV;
                        _IOdata.PsLo = do_cfgs.DOCfg[i].PsLo;
                        _IOdata.PsHi = do_cfgs.DOCfg[i].PsHi;
                        _IOdata.HDT = do_cfgs.DOCfg[i].HDT;
                        _IOdata.LDT = do_cfgs.DOCfg[i].LDT;
                        _IOdata.ACh = do_cfgs.DOCfg[i].ACh;
                        _IOdata.AMd = do_cfgs.DOCfg[i].AMd;
                        _IOdata.Tag = do_cfgs.DOCfg[i].Tag;
                        i++;
                    }
                }
                else
                    throw new Exception("Parser DO Data Fail");
            }
            catch (Exception e)
            {
                OnGetIOConfigHttpRequestError(e);
            }
            finally
            {
                System.GC.Collect();
            }
        }
        #endregion
        //------------------------------------------------------------------------//




    }

    public enum ServiceAction
    {
        Idel = 0,
        Connection = 1,
        GetDevData = 2,
        GetModbusCoilsAddr = 3,
        GetModbusCoilsLength = 4,
        GetModbusRegsAddr = 5,
        GetModbusRegsLength = 6,
        GetDIVal = 7,
        GetDOVal = 8,
        GetAIVal = 9,
        GetAOVal = 10,
        GetUInMod = 11,

        UpdateFW = 19,
        Done = 20,
    }    
}
