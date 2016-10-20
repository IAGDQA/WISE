namespace Model
{
    /////////////////////////////--------------------------------------------------------------
    /// <summary>
    ///emun WISE series device
    /// </summary>
    public enum WISEType
    {
        WISE4010LAN = 4010,
        WISE4050LAN = 4050,
        WISE4060LAN = 4060,
        WISE4012 = 24012,
        WISE4012E = 34012,
        WISE4050 = 24050,
        WISE4051 = 24051,
        WISE4060 = 24060,
    }
    /// <summary>
    ///emun Http request method
    /// </summary>
    public enum HttpRequestOption
    {
        None,
        GET,
        POST,
        PATCH,
    }
    /// <summary>
    ///emun WISE series Restful URI 
    /// </summary>
    public enum WISE_RESTFUL_URI
    {
        None,
        config,
        file_xfer,
        profile,
        sys_info,
        gen_config,
        control,
        net_basic,
        net_config,
        di_config,
        di_value,
        do_config,
        do_value,
        ai_genconfig,
        ai_config,
        ai_value,
        uio_config,
        modbus_coilconfig,
        modbus_regconfig,
        modbus_coilbas,
        modbus_coillen,
        modbus_regbas,
        modbus_reglen,
        accessctrl,
        log_dataoption,
        log_control,
        log_output,
        log_message,
        log_event,
        logex_output,
        logex_file,
        logex_list,
        datastream,
        p2p_mode,
        p2p_basic,
        p2p_advanced,
        wlan_config,
        wlan_value,
        logsys_output,
        logsys_message,
        log_upload,
        logsys_dataoption,
        log_switch,
        cloud_config,
        cloud_value,
        serial_config,
        modbusslave_genconfig,
    }

    public enum ValueRange
    {
        V_OMIT = -1,
        V_Neg10To10 = 323,
        V_Neg5To5 = 322,
        V_Neg2pt5To2pt5 = 321,
        V_Neg1To1 = 320,
        V_0To10 = 328,
        V_0To5 = 327,
        V_0To1 = 325,

        mV_Neg500To500 = 260,
        mV_Neg150To150 = 259,

        mV_0To500 = 262,
        mV_0To150 = 261,
        mA_Neg20To20 = 385,
        mA_0To20 = 386,
        mA_4To20 = 384,
    }
}
