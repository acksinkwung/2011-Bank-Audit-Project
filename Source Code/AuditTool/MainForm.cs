/* 
 * 2011 年 10 月 17 日
 * 
 * 稽核小工具
 * ver 1.4
 * 
 * 葉心寬 
 * Leo Yeh
 * 
 */

using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using System.Net.Sockets;
using System.Net;
using System.IO;
using System.Threading;
using System.Xml;
using XBRL;
using Newtonsoft.Json.Linq;

namespace AuditTool
{

    public partial class MainForm : Form
    {
        // 相關的變數宣告

        const int MAX_FILE_NUMBER = 65535;
        const int MAX_ERROR_NUMBER = 1024;
        const int MAX_HOST_NUMBER = 128;

        handleClinet[] client;

        TcpListener serverSocket;
        TcpClient clientSocket = null;
        NetworkStream serverStream = null;
        NetworkStream clientStream = null;

        static string[] errorServer = new string[MAX_ERROR_NUMBER];
        static string[] errorVersionRoot = new string[MAX_ERROR_NUMBER];
        static string[] errorOnlineRoot = new string[MAX_ERROR_NUMBER];
        static string[] errorFileName = new string[MAX_ERROR_NUMBER];
        static string[] errorStatus = new string[MAX_ERROR_NUMBER];

        static Dictionary<string, int> dtValue;
        static DataTable[] dtHost;
        static int countError;
        static DataTable dtError = new DataTable();
        static ExcelExporter ee = new ExcelExporter();
        static List<string> messageList;

        static Thread ctThread ;
        static Thread thread;
        public MainForm()
        {
            InitializeComponent();
        }
        /// <summary>
        /// 當按下"啟動伺服器"按鈕的動作
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnServer_Click(object sender, EventArgs e)
        {
            try
            {
                txtMessage.Text += " [ " + DateTime.Now + " ] 啟動伺服器" + Environment.NewLine;
                serverSocket = new TcpListener(IPAddress.Parse(getServerIP()), getServerPort());
                clientSocket = default(TcpClient);
                serverSocket.Start();
                client = new handleClinet[MAX_HOST_NUMBER];
                thread = new Thread(run);
                thread.Start();
                // 需啟動伺服器成功後才能進行稽核
                btnAudit.Enabled = true;
                txtMessage.Text += " [ " + DateTime.Now + " ] 伺服器啟動成功" + Environment.NewLine;
            }
            catch
            {
                txtMessage.Text += " [ " + DateTime.Now + " ] 伺服器啟動失敗" + Environment.NewLine;
            }
        }
        /// <summary>
        /// 當按下"進行稽核"按鈕的動作
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnAudit_Click(object sender, EventArgs e)
        {

                txtMessage.Text += " [ " + DateTime.Now + " ] 開始進行稽核" + Environment.NewLine;
                dtError.Clear();
                for (int n = 0; n < MAX_ERROR_NUMBER; n++)
                {
                    dtError.Rows.Add(errorServer[n], errorVersionRoot[n], errorOnlineRoot[n], errorFileName[n], errorStatus[n]);
                    if (errorServer[n] == null)
                    {
                        break;
                    }

                }

                ee = new ExcelExporter();
                ee.Add("有問題的檔案", dtError);
                foreach (KeyValuePair<string, int> value in dtValue)
                {
                    txtMessage.Text += " [ " + DateTime.Now + " ] 產生 [ " + value.Key + "] 的稽核結果" + Environment.NewLine;
                    ee.Add(value.Key, dtHost[value.Value]);
                }
                try
                {
                    string fileName = "Audit Result " + DateTime.Today.Year.ToString("D4") + DateTime.Today.Month.ToString("D2") + DateTime.Today.Day.ToString("D2") + ".xlsx";
                    ee.ExportDataTable(fileName);
                    txtMessage.Text += " [ " + DateTime.Now + " ] 稽核工作完成" + Environment.NewLine;
                    txtMessage.Text += " [ " + DateTime.Now + " ] 已產生檔案 " + fileName + Environment.NewLine;
                }
                catch {
                    txtMessage.Text += " [ " + DateTime.Now + " ] 檔案已開啟，無法進行儲存" + Environment.NewLine;
                }
        }
        /// <summary>
        /// 當按下"傳輸資料"按鈕的動作
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnClient_Click(object sender, EventArgs e)
        {
            try
            {
                clientSocket = new TcpClient();
                clientSocket.Connect(getServerIP(), getServerPort());
                clientStream = clientSocket.GetStream();
                txtMessage.Text += " [ " + DateTime.Now + " ] 傳送稽核相關資料至伺服器" + Environment.NewLine;


                UnicodeEncoding unien = new UnicodeEncoding();
                string clHostName = System.Windows.Forms.SystemInformation.ComputerName;
                string versionPath = getComparePathConfig(clHostName, "versionRoot", 0);
                string vrOutput = getDIRContent(versionPath, "") + "";
                
                string output = "{\"hn\":\"" + clHostName + "\",\"vr\":\"" +
                    System.Web.HttpUtility.UrlEncode(vrOutput) + "\",\"or\":";
                output += "[";
                int onlinePathCount = getCompareOnlinePathCount(clHostName);
                for (int n = 0; n < onlinePathCount; n++)
                {
                    output += "\"";
                    try
                    {
                        string onlinePath = getComparePathConfig(clHostName, "onlineRoot", n);
                        string orOutput = getDIRContent(onlinePath, "") + "";
                        output += System.Web.HttpUtility.UrlEncode(onlinePath+"|"+orOutput);
                    }
                    catch
                    {
                    }
                    output += "\"";
                    if (n != onlinePathCount - 1)
                    {
                        output += ",";
                    }
                    
                }

                output += "]}";
                clientSocket.SendBufferSize = unien.GetBytes(output).Length * 2;
                clientStream.Write(unien.GetBytes(output), 0, unien.GetBytes(output).Length);
                /*   
                int offset = 0;
                int sendSize = 65534;
                while (offset < output.Length)
                {
                    serverStream.Write(unien.GetBytes(output.Substring(offset, sendSize)), 0, unien.GetBytes(output.Substring(offset,sendSize)).Length);
                    offset += sendSize;
                }*/
                btnClient.Enabled = false;
            }
            catch
            {
                txtMessage.Text += " [ " + DateTime.Now + " ] 伺服器未開啟" + Environment.NewLine;
            }

        }
        /// <summary>
        /// 等待端點的連線
        /// </summary>
        private void run()
        {
            int counter = 0;
            while (true)
            {
                try
                {
                    clientSocket = serverSocket.AcceptTcpClient();
                    client[counter] = new handleClinet();
                    client[counter].startClient(clientSocket, Convert.ToString(counter));
                    counter += 1;
                }
                catch
                {
                }
            }
        }
        /// <summary>
        /// 取得設定檔中所指定伺服器的 IP
        /// </summary>
        /// <returns>IP 位址</returns>
        public static string getServerIP()
        {
            string fullName = "config.xml";
            XmlDocument doc = new XmlDocument();
            doc.Load(fullName);
            XmlNode node = doc.SelectSingleNode("//server");
            return node.Attributes["ip"].Value;
        }
        /// <summary>
        /// 取得設定檔中所指定伺服器的連接 Port
        /// </summary>
        /// <returns>連接 Port</returns>
        public static int getServerPort()
        {
            string fullName = "config.xml";
            XmlDocument doc = new XmlDocument();
            doc.Load(fullName);
            XmlNode node = doc.SelectSingleNode("//server");
            return Int16.Parse(node.Attributes["port"].Value);
        }
        /// <summary>
        /// 取得正式環式比較的路徑總數
        /// </summary>
        /// <param name="hostname"></param>
        /// <returns></returns>
        public static int getCompareOnlinePathCount(string hostname)
        {
            string fullName = "config.xml";
            XmlDocument doc = new XmlDocument();
            doc.Load(fullName);
            XmlNodeList nodes = doc.SelectNodes("//client");
            foreach (XmlNode node in nodes)
            {
                if (node.Attributes["name"].Value == hostname)
                {
                    return node["onlineRoot"].InnerText.Split('\n').Length-2;
                }
            }
            return 0;
        }
        /// <summary>
        /// 取得進行比較相關的路徑
        /// </summary>
        /// <param name="hostname">端點的名稱</param>
        /// <param name="type">類型(versionRoot : 版本環境，onlineRoot : 上線環境)</param>
        /// <returns></returns>
        public static string getComparePathConfig(string hostname, string type,int number)
        {
            string fullName = "config.xml";
            XmlDocument doc = new XmlDocument();
            doc.Load(fullName);
            XmlNodeList nodes = doc.SelectNodes("//client");
            foreach (XmlNode node in nodes)
            {
                if (node.Attributes["name"].Value == hostname)
                {
                    if (type == "versionRoot")
                    {
                        return (node["versionRoot"].InnerText.Replace("\r", "").Replace("\t", "").Split('\n'))[1];
                    }
                    if (type == "onlineRoot")
                    {
                        string[] onlinePath = node["onlineRoot"].InnerText.Replace("\r", "").Replace("\t", "").Split('\n');
                        return onlinePath[number+1];
                    }
                }
            }
            return "";
        }
        /// <summary>
        /// 取得目錄下的所有檔案和資料夾的資訊
        /// </summary>
        /// <param name="path">目錄路徑</param>
        /// <param name="level">子目錄</param>
        /// <returns></returns>
        public static string getDIRContent(string path,string level)
        {
            // 透過 dir /a 的指令取得所需的資訊
            Process p = new Process();
            p.StartInfo.UseShellExecute = false;
            p.StartInfo.RedirectStandardOutput = true;
            p.StartInfo.CreateNoWindow = true;
            p.StartInfo.FileName = "cmd";
            p.StartInfo.Arguments = "/c dir /a \"" + path + "\"";
            p.Start();
            string output = p.StandardOutput.ReadToEnd();
            p.WaitForExit();

            int c = 0, n = 0, f = 0;
            string result = "";

            // 分析 DIR 指令輸出的結果，進行排列組合，以行為單位
            char[] split = { '\r' };
            string[] lines = output.Split(split);

            foreach (string line in lines)
            {
                char[] splitLine = { ' ', '\n' };
                // 從第五行開始分析
                if (f >= 5)
                {
                    // 當是資料夾時，則進行該資料夾進行分析，透過遞迴的方式，回傳子目錄下的所有結果
                    if (line.IndexOf("<DIR>") > -1 )
                    {
                        if (line.IndexOf(".") == -1 && line.IndexOf("..") == -1)
                        {

                            string dataDIR = line.Substring(line.IndexOf("<DIR>") + 5, line.Length - (line.IndexOf("<DIR>")+5)).TrimEnd().TrimStart();

                            result += getDIRContent(path + "\\" + dataDIR, level+"\\"+dataDIR);
                        }
                        continue;
                    }
                    string[] datas = line.Split(splitLine);
                    c = 0;
                    foreach (string data in datas)
                    {
                        if (data != "")
                        {
                            c = c + 1;

                            if (c == 1)
                            {
                                // 檔案日期
                                if (data.IndexOf("<DIR>") > -1)
                                {
                                    c = c + 1;
                                }
                                else
                                {

                                    result += data + " ";
                                }
                            }
                            else if (c == 4)
                            {
                                // 檔案大小
                                result += data + " "; ;
                            }
                            else if (c == 5)
                            {
                                // 檔案名稱                                
                                result += level.Trim().Replace(" ","")+"\\"+data;


                            }
                            else if (c > 5)
                            {
                                // 檔案名稱                                
                                result += data.Trim().Replace(" ", "");
   
                            }
                        }
                    }


                }
                f = f + 1;
                // 每一筆檔案資訊，以行為單元
                result += "\n";

            }
            return result;
        }
        /// <summary>
        /// 處理端點的資訊
        /// </summary>
        class handleClinet
        {
            public TcpClient clientSocket;
            public string clHostName;
            public string clResult;
            public string clNo;
            /// <summary>
            /// 初始化
            /// </summary>
            /// <param name="inClientSocket"></param>
            /// <param name="clineNo"></param>
            public void startClient(TcpClient inClientSocket, string clineNo)
            {
                this.clientSocket = inClientSocket;
                this.clNo = clineNo;
                ctThread = new Thread(doGet);
                ctThread.Start();
            }
            /// <summary>
            /// 當接收到端點傳輸資料的動作
            /// </summary>
            private void doGet()
            {
                while (true)
                {
                    Stream stm = clientSocket.GetStream();
                    clientSocket.ReceiveBufferSize = 65536 * 2;
                    JObject o = null;
                    if (stm.CanRead)
                    {
                        clResult = "";
                        bool flag = true;
                        while (flag)
                        {
                            try
                            {
                                int k = 0;
                                byte[] bb = new byte[65536];

                                k = stm.Read(bb, 0, bb.Length);

                                if (k != 65536)
                                {
                                    flag = false;
                                }

                                clResult += System.Text.Encoding.Unicode.GetString(bb);
                            }
                            catch
                            {
                            }
                        }
                        try
                        {
                            o = JObject.Parse(clResult.Replace("\0", ""));
                            clHostName = o["hn"].ToString();
                        }
                        catch
                        {
                            recordMessage(" 傳輸失敗，請檢查網路是否有連線正常。 ");
                        }
                        string vrResult = "";
                        string[] orResult = new string[1024];

                        int vrFileNumber = 0;
                        int orFileNumber = 0;
                        dtValue.Add(clHostName, dtValue.Count + 1);

                        // 分析傳輸的資料，對應至相關的資訊
                        string[] name = new string[MAX_FILE_NUMBER];
                        string[] date = new string[MAX_FILE_NUMBER];
                        string[] size = new string[MAX_FILE_NUMBER];

                        string[] nameLocal = new string[MAX_FILE_NUMBER];
                        string[] dateLocal = new string[MAX_FILE_NUMBER];
                        string[] sizeLocal = new string[MAX_FILE_NUMBER];
                        string[] pathLocal = new string[MAX_FILE_NUMBER];

                        vrResult = System.Web.HttpUtility.UrlDecode(o["vr"].ToString());
                        string path = "";
                        int n = 0;
                        for (int i = 0; i < 1024; i++)
                        {
                            try
                            {
                                orResult[i] = System.Web.HttpUtility.UrlDecode(o["or"][i].ToString());
                                path = (orResult[i].Split('|'))[0];
                                orResult[i] = orResult[i].Replace(path + '|', "");
                            }
                            catch
                            {
                                break;
                            }
                            // 正式環境
                            try
                            {
                                string[] orResultDatas = orResult[i].Split('\n');

                                foreach (string orResultData in orResultDatas)
                                {
                                    string[] data = orResultData.Split(' ');
                                    try
                                    {
                                        if (data[2] != "")
                                        {
                                            dateLocal[n] = data[0];
                                            sizeLocal[n] = data[1];
                                            nameLocal[n] = data[2];
                                            pathLocal[n] = path;
                                            n++;
                                            if (nameLocal[n] != "")
                                            {
                                                orFileNumber = n;
                                            }
                                        }
                                    }
                                    catch
                                    {
                                    }
                                }
                            }
                            catch
                            {
                                break;
                            }
                        }


                        try
                        {


                            // 版本環境
                            string[] vrResultDatas = vrResult.Split('\n');
                            n = 0;
                            foreach (string vrResultData in vrResultDatas)
                            {
                                string[] data = vrResultData.Split(' ');
                                try
                                {
                                    date[n] = data[0];
                                    size[n] = data[1];
                                    name[n] = data[2];
                                    n++;
                                    if (name[n] != "")
                                    {
                                        vrFileNumber = n;
                                    }
                                }
                                catch
                                {
                                }
                            }

                            if (dtHost[dtValue[clHostName]].Columns.Count == 0)
                            {
                                dtHost[dtValue[clHostName]].Columns.Add("檔案名稱", typeof(string));
                                dtHost[dtValue[clHostName]].Columns.Add("版本環境日期", typeof(string));
                                dtHost[dtValue[clHostName]].Columns.Add("版本環境大小", typeof(string));
                                dtHost[dtValue[clHostName]].Columns.Add("正式環境日期", typeof(string));
                                dtHost[dtValue[clHostName]].Columns.Add("正式環境大小", typeof(string));
                            }
                            dtHost[dtValue[clHostName]].Clear();

                            for (int j = 0; j < orFileNumber; j++)
                            {

                                if (nameLocal[j] != null && pathLocal[j] + nameLocal[j] != pathLocal[j] && nameLocal[j].IndexOf("\0") == -1 && nameLocal[j] != "." && nameLocal[j] != "..")
                                {
                                    if (nameLocal[j].IndexOf(".") == -1)
                                    {
                                        break;
                                    }

                                    for (int i = 0; i < vrFileNumber; i++)
                                    {
                                        if (name[i] == nameLocal[j])
                                        {
                                            dtHost[dtValue[clHostName]].Rows.Add(pathLocal[j] + nameLocal[j], date[i], size[i], dateLocal[j], sizeLocal[j]);
                                            if (DateTime.Parse(date[i]) != DateTime.Parse(dateLocal[j]))
                                            {
                                                errorServer[countError] = clHostName;
                                                errorStatus[countError] = "檔案日期不一樣";
                                                errorVersionRoot[countError] = getComparePathConfig(clHostName, "versionRoot", 0);
                                                errorOnlineRoot[countError] = pathLocal[j];
                                                errorFileName[countError] = nameLocal[j];
                                                countError++;
                                            }
                                            if (size[i] != sizeLocal[j])
                                            {
                                                errorServer[countError] = clHostName;
                                                errorStatus[countError] = "檔案大小不一樣";
                                                errorVersionRoot[countError] = getComparePathConfig(clHostName, "versionRoot", 0);
                                                errorOnlineRoot[countError] = pathLocal[j];
                                                errorFileName[countError] = nameLocal[j];
                                                countError++;
                                            }
                                            break;
                                        }
                                        if (i == vrFileNumber - 1)
                                        {
                                            errorServer[countError] = clHostName;
                                            errorStatus[countError] = "可疑檔案被新增";
                                            errorVersionRoot[countError] = getComparePathConfig(clHostName, "versionRoot", 0);
                                            errorOnlineRoot[countError] = pathLocal[j];
                                            errorFileName[countError] = nameLocal[j];
                                            countError++;
                                            dtHost[dtValue[clHostName]].Rows.Add(pathLocal[j] + nameLocal[j], "不存在", "不存在", dateLocal[j], sizeLocal[j]);

                                        }
                                    }
                                }
                            }
                        }
                        catch
                        {
                        }
                        recordMessage(" [ " + DateTime.Now + " ] 電腦 [ " + clHostName + " ] 完成相關檔案傳輸。");
                    }
                }
            }
            
        }
        /// <summary>
        /// 暫存記錄訊息等待顯示
        /// </summary>
        /// <param name="msg">記錄訊息</param>
        private static void recordMessage(string msg) {
            messageList.Add(msg);
        }
        /// <summary>
        /// 小工具初始化
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form1_Load(object sender, EventArgs e)
        {

            messageList = new List<string>();
            this.FormClosing += new FormClosingEventHandler(Form1_FormClosing);
            dtHost = new DataTable[MAX_HOST_NUMBER];
            for (int n = 0; n < dtHost.Length; n++)
            {
                dtHost[n] = new DataTable();
            }
            dtValue = new Dictionary<string, int>();
            countError = 0;
            dtError = new DataTable();
            
            dtError.Columns.Add("伺服器", typeof(string));
            dtError.Columns.Add("版本環境根目錄", typeof(string));
            dtError.Columns.Add("正式環境根目錄", typeof(string));
            dtError.Columns.Add("檔案名稱", typeof(string));
            dtError.Columns.Add("狀態", typeof(string));
            btnAudit.Enabled = false;
            
        }
        
        /// <summary>
        /// 更新訊息的動作
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void checkTimer_Tick(object sender, EventArgs e)
        {
            for (int n = 0; n < messageList.Count; n++)
            {
                txtMessage.Text += messageList[n] + Environment.NewLine; ;
            }
            messageList.Clear();
        }
        /// <summary>
        /// 當關閉或結束程式時的動作
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
                try { 
                ctThread.Abort();
                thread.Abort();
                clientSocket.Close();
                }catch{

                }
                for (int n = 0; n < MAX_HOST_NUMBER; n++)
                {
                    try
                    {
                        client[n].clientSocket.Close();
                    }
                    catch
                    {
                    }
                }
                try
                {
                    serverSocket.Stop();
                }
                catch
                {
                }
        }
    }
}
