namespace Model
{
    /// <summary>
    /// Enum Device type
    /// </summary>
    public enum DevBase
    {
        APAX_Ctrllor = 1,
        WISE40XX = 2,
        APAX5070 = 3,
        ModbusTCP = 4,
    }

    public enum DevIDDefine
    {
        DI = 0,
        DO = 20,//offset 20
        AI = 40,//offset 40
        AO = 60,//offset 60
    }

    public enum ExeCaseRes//test case result
    {
        None,
        Pass,
        Fail,
    }
}
