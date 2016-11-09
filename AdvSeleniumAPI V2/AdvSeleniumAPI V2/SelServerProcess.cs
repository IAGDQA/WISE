using System;
using System.Diagnostics;

namespace AdvSeleniumAPI
{
    public class ServerProcess
    {
        int nodeId = 0;
        public ServerProcess()
        {
            //foreach (Process tp in Process.GetProcessesByName("cmd"))
            //{
            //    tp.CloseMainWindow();
            //    tp.Kill();
            //}
            //Process proc = null;
            //try
            //{
            //    proc = new Process();
            //    proc.StartInfo.WorkingDirectory = @"c:\";
            //    proc.StartInfo.FileName = "seleniumServerBat.bat";
            //    proc.StartInfo.CreateNoWindow = false;
            //    proc.Start();
            //    //Id = proc.Id;
            //    //proc.WaitForExit();
            //}
            //catch (Exception ex)
            //{
            //    //Console.WriteLine("Exception Occurred :{0},{1}", ex.Message, ex.StackTrace.ToString());
            //}


        }
        public void OpenServer()
        {
            
        }
        public void OpenNode()
        {
            Process proc = null;
            try
            {
                proc = new Process();
                proc.StartInfo.WorkingDirectory = @"c:\";
                proc.StartInfo.FileName = "seleniumServerNodeBat.bat";
                proc.StartInfo.CreateNoWindow = false;
                proc.Start();
                nodeId = proc.Id;
                proc.WaitForExit(3000);
            }
            catch (Exception ex)
            {
                //Console.WriteLine("Exception Occurred :{0},{1}", ex.Message, ex.StackTrace.ToString());
            }
            //Process p = new Process();
            //p.StartInfo.FileName = "cmd.exe";
            //p.StartInfo.WorkingDirectory = @"c:\";
            //p.StartInfo.UseShellExecute = false;
            //p.StartInfo.RedirectStandardOutput = true;
            //p.StartInfo.RedirectStandardInput = true;
            //p.Start();
            //p.StandardInput.WriteLine(@"java -jar selenium-server-standalone-2.53.1.jar -role node  -hub http://localhost:4444/grid/register");
            //nodeId = p.Id;
            ////p.WaitForExit(3000);
            //System.Threading.Thread.Sleep(3000);
        }
        public void CloseNode()
        {
            if (nodeId == 0) return;
            try
            {
                Process proKill = Process.GetProcessById(nodeId);
                proKill.CloseMainWindow();
                proKill.Kill();
            }
            catch (Exception ex)
            {
                //Console.WriteLine("Exception Occurred :{0},{1}", ex.Message, ex.StackTrace.ToString());
            }
        }

    }//class
}
