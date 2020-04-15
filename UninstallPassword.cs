using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Win32;
using System.IO;
using System.Reflection;
using System.Diagnostics;
using System.Data;
using System.Threading;
using System.Net;
//using Newtonsoft.Json.Linq;
using System.Management;
using System.Windows.Forms;

namespace eXrepUninstall
{
    public class UninstallPassword
    {
        public String sUninstallPassword = "eXsoft222";
        public String sResourcePath = "";
        public String sCoID = "";
        public String sUserID = "";
        public String sVersion = "";
        public static Dictionary<string, string> _CodeConfig = null;
        public static ReaderWriterLock rwl = new ReaderWriterLock();
        
		// 뭔가 수정을 했음... 뭔가... 2222
		
		// 슬랙테스트용 커밋..

        // 암호를 입력해야하는 사용자인지 확인
        public bool PasswordCheck()
        {
            bool bPWcheck = true;
            String sResourcePath = getAssemblyResourcePath();
            String sRegistryCoKey = GetConfigValue(sResourcePath, "RegistryCoKey");
            String sRegistryProductKey = GetConfigValue(sResourcePath, "RegistryProductKey");
            String sUninstallPasswordKey = GetConfigValue(sResourcePath, "uninstallPassword");

            sUninstallPassword = getUninstallPassword(sRegistryCoKey, sRegistryProductKey, sUninstallPasswordKey);

            if (sUninstallPassword == null || sUninstallPassword.Equals(""))
                bPWcheck = false;

            return bPWcheck;
        }


        // 암호 가져옴
        public string get_Password()
        {
            String sResourcePath = getAssemblyResourcePath();
            String sRegistryCoKey = GetConfigValue(sResourcePath, "RegistryCoKey");
            String sRegistryProductKey = GetConfigValue(sResourcePath, "RegistryProductKey");
            String sUninstallPasswordKey = GetConfigValue(sResourcePath, "uninstallPassword");

            sUninstallPassword = getUninstallPassword(sRegistryCoKey, sRegistryProductKey, sUninstallPasswordKey);

            if (sUninstallPassword == null || sUninstallPassword.Equals(""))
                return sUninstallPassword;

            return sUninstallPassword;
        }


        public void post_Uninstall()
        {
            String sRegistryCoKey = GetConfigValue(getAssemblyResourcePath(), "RegistryCoKey");
            String sRegistryProductKey = GetConfigValue(getAssemblyResourcePath(), "RegistryProductKey");
            String sRegistryUserKey = GetConfigValue(getAssemblyResourcePath(), "loginUserKey");

            String sUserID = getCurrentUserRegistry(sRegistryCoKey, sRegistryProductKey, sRegistryUserKey);
            if (sUserID != null && !sUserID.Equals(""))
            {
                //UpdateInstalledVersionAndConnectInfo("UNINSTALLED");
            }
        }


        public void OnAfterUninstall()
        {
            //restart_Explorer();
            removeUSBReadonly();
            // 가상 드라이브 해제 및 삭제
            if (detachVDisk())
            {
                //eXrepExplorerUtil.CommonUtil.deleteVDiskFile();
            }

            //unRegistryOutlookAdd();                                       // Setup Project의 설정으로 이동 됨
            setLocalControl(false);
        }


        // =======


        public static string getAssemblyResourcePath()
        {
            String sRet = new DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData)).FullName;

            Assembly assembly = Assembly.GetExecutingAssembly();
            //MessageBox.Show(assembly.ToString() + "    //    " + assembly.Location.ToString());

            //string dllPath = Directory.GetCurrentDirectory() + @"\eXrepUninstall.dll";
            //string dllPath = Environment.CurrentDirectory + @"\eXrepUninstall.dll";

            // eXrepExplorer.dll의 CompanyName, ProductName 참조하도록  
            string dllPath = Environment.CurrentDirectory + @"\eXrepExplorer.dll";

            FileVersionInfo fvi = FileVersionInfo.GetVersionInfo(dllPath); 

            String companyName = fvi.CompanyName;
            String productNAme = fvi.ProductName;

            sRet += "\\" + companyName + "\\" + productNAme;

            return sRet;
        }


        public String getUninstallPassword(string sFirstKey, string sSecondKey, String sValueKey)
        {
            String sGetValue = "";
            RegistryKey regKeyAppRoot = null;
            UTF8Encoding utf8 = new UTF8Encoding();

            try
            {
                regKeyAppRoot = Registry.LocalMachine.OpenSubKey("SOFTWARE\\" + sFirstKey, true);
                if (regKeyAppRoot != null)
                {
                    using (RegistryKey regKeySub = regKeyAppRoot.OpenSubKey(sSecondKey, true))
                    {
                        if (regKeySub != null)
                        {
                            object oRegData = regKeySub.GetValue(sValueKey);
                            if (oRegData != null)
                            {
                                sGetValue = oRegData.ToString();

                                Byte[] getBytes = Convert.FromBase64String(sGetValue);
                                sGetValue = utf8.GetString(getBytes);
                            }
                        }
                    }
                }

                if (sGetValue == null || sGetValue.Equals(""))
                {
                    if (regKeyAppRoot != null) regKeyAppRoot.Close();

                    regKeyAppRoot = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Wow6432Node\\" + sFirstKey, true);
                    if (regKeyAppRoot != null)
                    {
                        using (RegistryKey regKeySub = regKeyAppRoot.OpenSubKey(sSecondKey, true))
                        {
                            if (regKeySub != null)
                            {
                                object oRegData = regKeySub.GetValue(sValueKey);
                                if (oRegData != null)
                                {
                                    sGetValue = oRegData.ToString();

                                    Byte[] getBytes = Convert.FromBase64String(sGetValue);
                                    sGetValue = utf8.GetString(getBytes);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception)
            {
            }
            finally
            {
                if (regKeyAppRoot != null) regKeyAppRoot.Close();
            }

            return sGetValue;
        }


        public static string GetConfigValue(string sFilePath, string configCode)
        {
            string strTemp = string.Empty;

            try
            {
                if (_CodeConfig == null)
                {
                    rwl.AcquireWriterLock(500); //쓰기Lock

                    try
                    {
                        GetCodeConfigList(sFilePath);
                        //bFinishedToRead = true;
                    }
                    catch
                    {
                        _CodeConfig = null;
                        //bFinishedToRead = false;
                        throw;
                    }
                    finally
                    {
                        rwl.ReleaseWriterLock();
                    }
                }

                rwl.AcquireReaderLock(10000); //읽기Lock

                try
                {
                    if (!_CodeConfig.ContainsKey(configCode))
                    {
                        strTemp = configCode;
                    }
                    else
                    {
                        string entry = _CodeConfig[configCode];

                        strTemp = entry != string.Empty ? entry : configCode;
                    }
                }
                finally
                {
                    rwl.ReleaseReaderLock();
                }
            }
            catch (Exception)
            {
                //MessageBox.Show(ex.ToString(), "Cloud Signature", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return strTemp;
        }


        public static void GetCodeConfigList(string sFilePath)
        {
            try
            {
                //Dictionary항목을 담는다.
                _CodeConfig = new Dictionary<string, string>();

                string _Code = string.Empty;
                string _Value = string.Empty;
                string _XmlFile = string.Empty;

                _XmlFile = sFilePath + @"\CommonProp.xml";

                using (DataSet oDS = new DataSet())
                {
                    oDS.ReadXml(_XmlFile);

                    for (int _tCnt = 0; _tCnt < oDS.Tables.Count; _tCnt++)
                    {
                        for (int _Cnt = 0; _Cnt < oDS.Tables[_tCnt].Columns.Count; _Cnt++)
                        {
                            try
                            {
                                _Code = oDS.Tables[_tCnt].Columns[_Cnt].ColumnName.ToString().Trim();
                                _Value = oDS.Tables[_tCnt].Rows[0][_Cnt].ToString().Trim();

                                _CodeConfig.Add(_Code, _Value);
                            }
                            catch (Exception) { }
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


        //================= post_Uninstall


        public static String getCurrentUserRegistry(string sFirstKey, string sSecondKey, String sValueKey, string subKey = "")
        {
            String sGetValue = "";
            RegistryKey regKeyAppRoot = null;
            UTF8Encoding utf8 = new UTF8Encoding();

            try
            {
                regKeyAppRoot = Registry.CurrentUser.OpenSubKey("SOFTWARE\\" + sFirstKey, true);
                if (regKeyAppRoot != null)
                {
                    if (subKey.Equals(""))
                    {
                        using (RegistryKey regKeySub = regKeyAppRoot.OpenSubKey(sSecondKey, true))
                        {
                            if (regKeySub != null)
                            {
                                object oRegData = regKeySub.GetValue(sValueKey);
                                if (oRegData != null)
                                {
                                    sGetValue = oRegData.ToString();

                                    Byte[] getBytes = Convert.FromBase64String(sGetValue);
                                    sGetValue = utf8.GetString(getBytes);
                                }
                            }
                        }
                    }
                    else 
                    {
                        subKey = sSecondKey + "\\" + subKey;
                        using (RegistryKey regKeySub = regKeyAppRoot.OpenSubKey(subKey, true))
                        {
                            if (regKeySub != null)
                            {
                                object oRegData = regKeySub.GetValue(sValueKey);
                                if (oRegData != null)
                                {
                                    sGetValue = oRegData.ToString();

                                    Byte[] getBytes = Convert.FromBase64String(sGetValue);
                                    sGetValue = utf8.GetString(getBytes);
                                }
                            }
                        }
                    }
                }

                if (sGetValue == null || sGetValue.Equals(""))
                {
                    if (regKeyAppRoot != null) regKeyAppRoot.Close();

                    regKeyAppRoot = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Wow6432Node\\" + sFirstKey, true);
                    if (regKeyAppRoot != null)
                    {
                        using (RegistryKey regKeySub = regKeyAppRoot.OpenSubKey(sSecondKey, true))
                        {
                            if (regKeySub != null)
                            {
                                object oRegData = regKeySub.GetValue(sValueKey);
                                if (oRegData != null)
                                {
                                    sGetValue = oRegData.ToString();

                                    Byte[] getBytes = Convert.FromBase64String(sGetValue);
                                    sGetValue = utf8.GetString(getBytes);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception)
            {
            }
            finally
            {
                if (regKeyAppRoot != null) regKeyAppRoot.Close();
            }

            return sGetValue;
        }



        public Boolean UpdateInstalledVersionAndConnectInfo(String sJob)
        {
            sResourcePath = getAssemblyResourcePath();  // 임의로 추가

            Boolean bRet = false;

            // Assign values to these objects here so that they can
            // be referenced in the finally block
            WebResponse response = null;
            StreamReader readerStream = null;

            // Use a try/catch/finally block as both the WebRequest and Stream
            // classes throw exceptions upon error
            try
            {
                // Create a request for the specified remote file name
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(GetConfigValue(sResourcePath, "InstalledVersionURL"));
                if (request != null)
                {
                    //string postData = "user_id=" + HttpUtility.UrlEncode(sUserID);
                    //postData += ("&folder_id=" + HttpUtility.UrlEncode(sFolderID));
                    //postData += ("&session_check=false");

                    //byte[] header = Encoding.UTF8.GetBytes(postData.ToString());
                    //long contentLength = header.Length;

                    StringBuilder sb = new StringBuilder();
                    sb.Append("{coId: \"" + sCoID + "\",");
                    sb.Append("authId: \"" + sUserID + "\",");
                    sb.Append("ipAddress: \"" + getClientIPAddress() + "\",");
                    sb.Append("installJob : \"" + sJob + "\",");
                    sb.Append("version : \"" + sVersion + "\"}");

                    byte[] sBody = Encoding.UTF8.GetBytes(sb.ToString());
                    long contentLength = sBody.Length;

                    request.Timeout = 120000;
                    request.Method = "POST";
                    request.ContentType = "application/json";
                    request.ContentLength = contentLength;

                    using (Stream reqStream = request.GetRequestStream())
                    {
                        reqStream.Write(sBody, 0, sBody.Length);
                    }

                    // Send the request to the server and retrieve the
                    // WebResponse object 
                    response = request.GetResponse();
                    if (response != null)
                    {
                        readerStream = new StreamReader(response.GetResponseStream());
                        string responseString = readerStream.ReadToEnd();
                        if (responseString != null && !responseString.Equals(""))
                        {
                            /*
                            JObject json = JObject.Parse(responseString);
                            string result = (string)json["status"];
                            if (result == "SUCCESS")
                            {
                                bRet = true;
                            }
                            */

                            int ResultNum = responseString.IndexOf("status") + 8;   // "status":"SUCCESS" -> SUCCESS 앞의 " 위치
                            if (responseString.Substring(ResultNum, 9).Contains("SUCCESS"))
                                bRet = true;
                        }
                    }
                }
            }
            catch (WebException)
            {
            }
            catch (UriFormatException)
            {
            }
            catch (Exception)
            {
            }
            finally
            {
                // Close the response and streams objects here 
                // to make sure they're closed even if an exception
                // is thrown at some point
                if (response != null) response.Close();
            }

            return bRet;
        }


        // 사용자 컴퓨터의 ip Address를 가져옴
        public static String getClientIPAddress()
        {

            String sRet = "";

            foreach (IPAddress IPA in Dns.GetHostAddresses(System.Net.Dns.GetHostName()))
            {
                if (IPA.AddressFamily.ToString() == "InterNetwork")
                {
                    sRet = IPA.ToString();
                    break;
                }
            }
            return sRet;
        }


        //================= OnAfterUninstall  


        public static void removeUSBReadonly()
        {
            RegistryKey regKeyAppRoot = null;

            try
            {
                regKeyAppRoot = Registry.LocalMachine.OpenSubKey(@"SYSTEM\CurrentControlSet\Control\StorageDevicePolicies", true);
                if (regKeyAppRoot == null) regKeyAppRoot = Registry.LocalMachine.CreateSubKey(@"SYSTEM\CurrentControlSet\Control\StorageDevicePolicies");

                if (regKeyAppRoot != null)
                {
                    regKeyAppRoot.SetValue("WriteProtect", 0, RegistryValueKind.DWord);
                }
            }
            catch (Exception) { }
            finally
            {
                if (regKeyAppRoot != null) regKeyAppRoot.Close();
            }
        }

        public static Boolean detachVDisk()
        {
            Boolean bRet = false;

            String sWindowPath = new DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.System)).FullName;
            String sVHDName = GetConfigValue(getAssemblyResourcePath(), "VDiskFileName");
            String sVDiskLetter = GetConfigValue(getAssemblyResourcePath(), "VDiskLetter");

            String sTakeoutVHDName = GetConfigValue(getAssemblyResourcePath(), "TakeoutDiskFileName");
            String sTakeoutDiskLetter = GetConfigValue(getAssemblyResourcePath(), "TakeoutDiskLetter");

            String sRegistryCoKey = GetConfigValue(getAssemblyResourcePath(), "RegistryCoKey");
            String sRegistryProductKey = GetConfigValue(getAssemblyResourcePath(), "RegistryProductKey");
            String sRegistryUserKey = GetConfigValue(getAssemblyResourcePath(), "loginUserKey");
            String sUserID = getCurrentUserRegistry(sRegistryCoKey, sRegistryProductKey, sRegistryUserKey);

            string VHDDir = sWindowPath;
            if (sRegistryProductKey.Equals("CS") && !sUserID.Equals(""))
            {
                VHDDir = getAssemblyResourcePath() + "\\" + sUserID;

                sVDiskLetter = getCurrentUserRegistry(sRegistryCoKey, sRegistryProductKey, "VISUALDISKLETTER", sUserID);
                sTakeoutDiskLetter = getCurrentUserRegistry(sRegistryCoKey, sRegistryProductKey, "TAKEOUTDISKLETTER", sUserID);
            }

            try
            {
                if (sVDiskLetter != null && !sVDiskLetter.Equals("") && Directory.Exists(sVDiskLetter + ":\\"))
                {
                    Process cmd = new Process();

                    ProcessStartInfo info = new ProcessStartInfo(sWindowPath + @"\diskpart.exe", "");
                    info.WorkingDirectory = sWindowPath;
                    //info.RedirectStandardInput = true;
                    info.RedirectStandardOutput = true;
                    info.RedirectStandardError = true;
                    info.UseShellExecute = false;
                    info.CreateNoWindow = true;
                    info.RedirectStandardInput = true;
                    info.Verb = "runas";
                    cmd.StartInfo = info;
                    cmd.Start();

                    //cmd.StandardInput.WriteLine("select vdisk file=\"" + sWindowPath + "\\" + sVHDName + "\""); 
                    cmd.StandardInput.WriteLine("select vdisk file=\"" + VHDDir + "\\" + sVHDName + "\"");
                    cmd.StandardInput.WriteLine("detach vdisk");
                    cmd.StandardInput.WriteLine("exit");
                    string output = cmd.StandardOutput.ReadToEnd();
                    cmd.WaitForExit();
                    cmd.Close();

                    bRet = true;
                }
            }
            catch (Exception)
            {
            }

            try
            {
                if (sTakeoutDiskLetter != null && !sTakeoutDiskLetter.Equals("") && Directory.Exists(sTakeoutDiskLetter + ":\\"))
                {
                    Process cmd = new Process();

                    ProcessStartInfo info = new ProcessStartInfo(sWindowPath + @"\diskpart.exe", "");
                    info.WorkingDirectory = sWindowPath;
                    //info.RedirectStandardInput = true;
                    info.RedirectStandardOutput = true;
                    info.RedirectStandardError = true;
                    info.UseShellExecute = false;
                    info.CreateNoWindow = true;
                    info.RedirectStandardInput = true;
                    info.Verb = "runas";
                    cmd.StartInfo = info;
                    cmd.Start();

                    //cmd.StandardInput.WriteLine("select vdisk file=\"" + sWindowPath + "\\" + sTakeoutVHDName + "\"");
                    cmd.StandardInput.WriteLine("select vdisk file=\"" + VHDDir + "\\" + sTakeoutVHDName + "\"");
                    cmd.StandardInput.WriteLine("detach vdisk");
                    cmd.StandardInput.WriteLine("exit");
                    string output = cmd.StandardOutput.ReadToEnd();
                    cmd.WaitForExit();
                    cmd.Close();

                    bRet = true;
                }
            }
            catch (Exception) { }

            return bRet;
        }

        public static void setLocalControl(Boolean isAutoupload)
        {
            //-------------------------------------------------------------- 로컬 드라이버 ---------------------------------------------------------------
            int iSetValue = getLocalDriveInt();
            RegistryKey regKeySub = null;

            try
            {
                //RegistryPermission readPerm = new RegistryPermission(RegistryPermissionAccess.AllAccess, @"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer");

                regKeySub = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer", true);
                if (regKeySub == null)
                {
                    regKeySub = Registry.LocalMachine.CreateSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer");
                    regKeySub.Close();
                }

                using (regKeySub = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer", true))
                {
                    if (regKeySub != null)
                    {
                        if (isAutoupload)
                        {
                            regKeySub.SetValue("NoDrives", iSetValue, RegistryValueKind.DWord);
                            regKeySub.SetValue("NoViewOnDrive", iSetValue, RegistryValueKind.DWord);
                        }
                        else
                        {
                            regKeySub.SetValue("NoDrives", 0, RegistryValueKind.DWord);
                            regKeySub.SetValue("NoViewOnDrive", 0, RegistryValueKind.DWord);
                        }
                    }
                }
            }
            catch (Exception) { }
            finally
            {
                if (regKeySub != null) regKeySub.Close();
            }
        }

        public static int getLocalDriveInt()
        {
            int iRet = 0;

            try
            {
                String sGetDrive = loadLocalDriver();
                if (sGetDrive != null && !sGetDrive.Equals(""))
                {
                    string[] sSplitDrive = sGetDrive.Split('*');
                    if (sSplitDrive.Length > 0)
                    {
                        for (int loop_count = 0; loop_count < sSplitDrive.Length; loop_count++)
                        {
                            String sDrive = sSplitDrive[loop_count];
                            if (sDrive != null && !sDrive.Equals(""))
                            {
                                if (sDrive.Equals("C:\\")) iRet += 4;
                                else if (sDrive.Equals("D:\\")) iRet += 8;
                                else if (sDrive.Equals("E:\\")) iRet += 16;
                                else if (sDrive.Equals("F:\\")) iRet += 32;
                                else if (sDrive.Equals("G:\\")) iRet += 64;
                                else if (sDrive.Equals("H:\\")) iRet += 128;
                                //else if (sDrive.Equals("I:\\")) iRet += 256;
                                //else if (sDrive.Equals("J:\\")) iRet += 512;
                                else if (sDrive.Equals("K:\\")) iRet += 1024;
                                else if (sDrive.Equals("L:\\")) iRet += 2048;
                                else if (sDrive.Equals("M:\\")) iRet += 4096;
                                else if (sDrive.Equals("N:\\")) iRet += 8192;
                                else if (sDrive.Equals("O:\\")) iRet += 16384;
                                else if (sDrive.Equals("P:\\")) iRet += 32768;
                                else if (sDrive.Equals("Q:\\")) iRet += 65536;
                                else if (sDrive.Equals("R:\\")) iRet += 131072;
                                else if (sDrive.Equals("S:\\")) iRet += 262144;
                                else if (sDrive.Equals("T:\\")) iRet += 524288;
                                else if (sDrive.Equals("U:\\")) iRet += 1048576;
                                else if (sDrive.Equals("V:\\")) iRet += 2097152;
                                else if (sDrive.Equals("W:\\")) iRet += 4194304;
                                else if (sDrive.Equals("X:\\")) iRet += 8388608;
                                else if (sDrive.Equals("Y:\\")) iRet += 16777216;
                                else if (sDrive.Equals("Z:\\")) iRet += 33554432;
                            }
                        }
                    }
                }
            }
            catch (Exception) { }

            return iRet;
        }

        public static String loadLocalDriver()
        {
            String sLocalDriver = "";

            try
            {
                ManagementObjectCollection drives = new ManagementObjectSearcher("SELECT Caption, DeviceID FROM Win32_DiskDrive").Get(); // WHERE InterfaceType<>'USB'
                foreach (ManagementObject drive in drives)
                {
                    foreach (ManagementObject partition in new ManagementObjectSearcher("ASSOCIATORS OF {Win32_DiskDrive.DeviceID='" + drive["DeviceID"] + "'} WHERE AssocClass = Win32_DiskDriveToDiskPartition").Get())
                    {
                        foreach (ManagementObject disk in new ManagementObjectSearcher("ASSOCIATORS OF {Win32_DiskPartition.DeviceID='" + partition["DeviceID"] + "'} WHERE AssocClass = Win32_LogicalDiskToPartition").Get())
                        {
                            if (sLocalDriver.Equals(""))
                                sLocalDriver = disk["CAPTION"].ToString() + "\\";
                            else
                                sLocalDriver += "*" + disk["CAPTION"].ToString() + "\\";
                        }
                    }
                }
            }
            catch (Exception) { }

            if (sLocalDriver == null || sLocalDriver.Equals("")) sLocalDriver = "C:\\";

            return sLocalDriver;
        }


    }
}
