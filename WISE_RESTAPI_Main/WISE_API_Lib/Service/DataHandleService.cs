﻿using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Runtime.InteropServices;

namespace Service
{
    public class DataHandleService
    {
        string filename = "WISE_API_CONFIG.ini";
        string folderPath = "";
        public DataHandleService()
        {

        }

        public string GetPara(string _path)
        {
            string resPath = "", folderPath = "";

            string sPath = System.Reflection.Assembly.GetAssembly(this.GetType()).Location;
            char delimiterChars = '\\';
            string[] words = sPath.Split(delimiterChars);

            for (int i = 0; i < words.Length - 1; i++)
            {
                folderPath = folderPath + words[i] + "\\";
            }

            if (File.Exists(folderPath + "\\" + filename))
            {
                using (ExecuteIniClass IniFile = new ExecuteIniClass(Path.Combine(folderPath, filename)))
                {
                    resPath = IniFile.getKeyValue("Dev", "IP");
                }
            }

            return resPath;
        }

        public bool SavePara(string _path, string _ip)
        {
            bool res = true;
            try
            {
                string sPath = System.Reflection.Assembly.GetAssembly(this.GetType()).Location;
                folderPath = Path.GetDirectoryName(sPath);
                //char delimiterChars = '\\';
                //string[] words = sPath.Split(delimiterChars);

                //for (int i = 0; i < words.Length - 1; i++)
                //{
                //    folderPath = folderPath + words[i] + "\\";
                //}
                //create new
                if (!File.Exists(folderPath + "\\" + filename))
                {
                    File.Create(folderPath + "\\" + filename);
                }

                //save para.
                using (ExecuteIniClass IniFile = new ExecuteIniClass(Path.Combine(folderPath, filename)))
                {
                    IniFile.setKeyValue("Dev", "IP", _ip);
                }
            }
            catch (Exception ex)
            {
                res = false;
            }

            return res;
        }

        public IniDataFmt GetGropPara(string _path)
        {
            IniDataFmt data = new IniDataFmt() { IP = "", Path = "", Browser = "" };
            folderPath = "";
            string sPath = System.Reflection.Assembly.GetAssembly(this.GetType()).Location;
            char delimiterChars = '\\';
            string[] words = sPath.Split(delimiterChars);

            for (int i = 0; i < words.Length - 1; i++)
            {
                folderPath = folderPath + words[i] + "\\";
            }

            if (File.Exists(folderPath + "\\" + filename))
            {
                using (ExecuteIniClass IniFile = new ExecuteIniClass(Path.Combine(folderPath, filename)))
                {
                    data.IP = IniFile.getKeyValue("Dev", "IP");
                    data.Path = IniFile.getKeyValue("Dev", "Path");
                    data.Browser = IniFile.getKeyValue("Dev", "Browser");
                }
            }

            return data;
        }

        public bool SaveGropPara(string _path, IniDataFmt _obj)
        {
            bool res = true;
            try
            {
                //string sPath = System.Reflection.Assembly.GetAssembly(this.GetType()).Location;
                //create new
                if (!File.Exists(folderPath + "\\" + filename))
                    File.Create(folderPath + "\\" + filename);

                //save para.
                using (ExecuteIniClass IniFile = new ExecuteIniClass(Path.Combine(folderPath, filename)))
                {
                    IniFile.setKeyValue("Dev", "IP", _obj.IP);
                    IniFile.setKeyValue("Dev", "Path", _obj.Path);
                    IniFile.setKeyValue("Dev", "Browser", _obj.Browser);
                }
            }
            catch (Exception ex)
            {
                res = false;
            }

            return res;
        }
    }
    public class IniDataFmt
    {
        public string IP { get; set; }
        public string Path { get; set; }
        public string Browser { get; set; }
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
