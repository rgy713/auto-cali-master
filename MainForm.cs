using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.IO;
using System.IO.Ports;
using Common;

namespace comTest
{
    public delegate void DelegatLog(string msg);
    public delegate void ViewStatusCallback(string msg);
    public partial class MainForm : Form
    {
        private string m_GZHComPort = null;
        private string m_CJBComPort = null;

        private int m_GZHBaudRate = 0;
        private int m_CJBBaudRate = 0;

        private SerialPort m_GZHSerialPort = null;
        private SerialPort m_CJBSerialPort = null;

        private Boolean m_GZHIsConnect = false;
        private Boolean m_CJBIsConnect = false;

        List<byte> m_GZHReceivedBytes = new List<byte>();
        List<byte> m_CJBReceivedBytes = new List<byte>();

        Thread m_GZHReadThread;
        Boolean m_IsGZHContinueRead = false;

        Thread m_CJBReadThread;
        Boolean m_IsCJBContinueRead = false;

        FileStream m_FileStream = null;
        StreamWriter m_StreamWriter = null;

        Export m_export = null;
        Export1 m_export1 = null;

        private float[][] m_OriginValues;
        private float[][][] m_Measurements;
        private float[][][] m_Errors;
        private float[][][] m_PhaseAngles;

        private string[] m_BodyStr = new string[] { "RL", "RA", "TT", "LL", "LA" };
        private string[] m_BodyAngleStr = new string[] { "RL", "RA", "TT", "LL", "LA", "RALL", "LARL"};
        private string[] m_RateStr = new string[] { "5k", "50k", "250k", "500k", "1k", "1000k" };
        private float m_ErrPerM1 = 1;
        private float m_ErrPerM2 = 3;
        private float[] m_TestWeightLevels;
        private float[][] m_TestWeightResults;
        private int m_SelectedTestWeight = -1;
        private string[] m_WeightPositionLevels;
        private int m_SelectedWeightPosition = 0;
        private string[] m_MachineTypeList;
        private string m_MachineType;
        private string m_SerialNumber;
        private string m_ProductSerialNumber;

        //记录阻抗数量
        private int m_Counter = 1;
        private int m_GZHCommandStep = 0;
        private int m_CJBCommandStep = 0;
        private int m_ControlStep = 0;
        private bool m_ForDevelopment = true;

        private byte m_HeaderByte = (byte)0x02;
        private byte m_FooterByte = (byte)0x03;
        private byte m_SecondByte = (byte)0x53;
        private Command m_CJBCommand;

        private Boolean m_IsAutoTest = true;
        //自动
        private Boolean m_chkIsOnlyTest = false;
        private Boolean m_chkIsOriginalR = false;
        private int m_cboSelectedTestItem = 0;

        //手动
        private int m_cbxSelectGZHPosition = 0;
        private Boolean m_IsOnlyTestSD = false;
        private int m_cboSelectedTestItemSD = 0;
        private Boolean m_IsOrgImpSD = false;
        private Boolean m_IsBodyTest = false;

        public MainForm()
        {
            InitializeComponent();

            m_CJBCommand = new Command(m_HeaderByte, m_SecondByte, m_FooterByte, 1);

            Loger.LogRecived += new DelegateLog(OnLog);
            Loger.LogSend += new DelegateLog(OnSendLog);
            try
            {
                m_FileStream = new FileStream("log.txt", FileMode.OpenOrCreate);
                m_StreamWriter = new StreamWriter(m_FileStream, Encoding.Unicode);
            }
            catch (Exception ex)
            {
                if (m_ForDevelopment) Loger.send(ex.ToString());
            }
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            Initionalization();
        }

        private void MainForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            try
            {
                CloseCJBSerialPort();

                Thread.Sleep(100);

                Params2Ini();

                m_StreamWriter.Close();
                m_FileStream.Close();
            }
            catch
            { }
        }

        private void cboxBaud2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cboxBaud2.Text != null && cboxBaud2.Text != "")
            {
                m_CJBBaudRate = Int32.Parse(cboxBaud2.Text);
            }
        }

        private void btnCJBOpen_Click(object sender, EventArgs e)
        {
            OpenCJBSerialPort();

            CJBOpenStatusButtons();
        }

        private void btnCJBClose_Click(object sender, EventArgs e)
        {
            CloseCJBSerialPort();

            InitButtons();
        }
        private void btnClearSendBox_Click(object sender, EventArgs e)
        {
            richboxSend.Clear();
        }

        private void btnClearReceiveBox_Click(object sender, EventArgs e)
        {
            richboxReceive.Clear();
        }

        private void btnInitStatus_Click(object sender, EventArgs e)
        {
            Initionalization();
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            m_export.m_MachineType = m_MachineType;
            m_export.m_SerialNumber = m_SerialNumber;
            m_export.m_ProductSerialNumber = m_ProductSerialNumber;
            m_export.m_Measurements = m_Measurements;
            m_export.m_Errors = m_Errors;
            m_export.m_TestWeightResults = m_TestWeightResults;
            m_export.m_PhaseAngles = m_PhaseAngles;

            m_export.export();
        }

        private void btnExport1_Click(object sender, EventArgs e)
        {
            m_export1.m_MachineType = m_MachineType;
            m_export1.m_SerialNumber = m_SerialNumber;
            m_export1.m_ProductSerialNumber = m_ProductSerialNumber;
            m_export1.m_Measurements = m_Measurements;
            m_export1.m_Errors = m_Errors;
            m_export1.m_TestWeightResults = m_TestWeightResults;
            m_export1.m_PhaseAngles = m_PhaseAngles;

            m_export1.export();
        }

        private void btnCalibration100_Click(object sender, EventArgs e)
        {
            if (CheckCJBCom())
            {
                btnCalibration100.BackColor = Color.Red;
                CalibrateBtnsEnable(false);
                SendCommand2CJB(100);
            }
        }

        private void btnCalibration75_Click(object sender, EventArgs e)
        {
            if (CheckCJBCom())
            {
                btnCalibration75.BackColor = Color.Red;
                CalibrateBtnsEnable(false);
                SendCommand2CJB(101);
            }
        }

        private void btnCalibration50_Click(object sender, EventArgs e)
        {
            if (CheckCJBCom())
            {
                btnCalibration50.BackColor = Color.Red;
                CalibrateBtnsEnable(false);
                SendCommand2CJB(102);
            }
        }

        private void btnCalibration25_Click(object sender, EventArgs e)
        {
            if (CheckCJBCom())
            {
                btnCalibration25.BackColor = Color.Red;
                CalibrateBtnsEnable(false);
                SendCommand2CJB(103);
            }
        }

        private void btnCalibration0_Click(object sender, EventArgs e)
        {
            if (CheckCJBCom())
            {
                btnCalibration0.BackColor = Color.Red;
                CalibrateBtnsEnable(false);
                SendCommand2CJB(104);
            }
        }

        private void btnGetWeight_Click(object sender, EventArgs e)
        {
            if (CheckCJBCom())
            {
                btnGetWeight.BackColor = Color.Red;
                txtWeight.Text = "";
                CalibrateBtnsEnable(false);
                SendCommand2CJB(105);
            }
        }
       
        
        private void cbxCom2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbxCom2.Text != null && cbxCom2.Text != "")
            {
                m_CJBComPort = cbxCom2.Text;
            }
        }

        
        private void cbxSelectGZHPosition_SelectedIndexChanged(object sender, EventArgs e)
        {
            m_cbxSelectGZHPosition = cbxSelectGZHPosition.SelectedIndex;
        }

        private void chkIsOrgImpSD_CheckedChanged(object sender, EventArgs e)
        {
            m_IsOrgImpSD = chkIsOrgImpSD.Checked;
        }

        private void btnSDTest_Click(object sender, EventArgs e)
        {
            m_IsAutoTest = false;
            btnSDTest.BackColor = Color.Red;
            CalibrateBtnsEnable(false);
            StartSDTest();
        }

        private void cbxTestWeight_SelectedIndexChanged(object sender, EventArgs e)
        {
            m_SelectedTestWeight = cbxTestWeight.SelectedIndex;
            cbxWeightPosition.SelectedIndex = 0;
            cbxWeightPosition.Enabled = (m_SelectedTestWeight == 1);
        }

        private void cbxWeightPosition_SelectedIndexChanged(object sender, EventArgs e)
        {
            m_SelectedWeightPosition = cbxWeightPosition.SelectedIndex;
        }

        private void cbxMachineType_SelectedIndexChanged(object sender, EventArgs e)
        {
            m_MachineType = cbxMachineType.Text;
            AutoCreateSerialNumber();
        }

        private void dupSerialNumber_SelectedItemChanged(object sender, EventArgs e)
        {
            m_ProductSerialNumber = dupSerialNumber.Text;

            m_SerialNumber = m_ProductSerialNumber.Substring(4, m_ProductSerialNumber.Length - 4);
        }
        private void btnBodyTest_Click(object sender, EventArgs e)
        {
            m_IsAutoTest = false;
            btnBodyTest.BackColor = Color.Red;
            CalibrateBtnsEnable(false);
            StartBodyTest();
        }

        private void SendLog(string msg)
        {
            richboxSend.AppendText(msg + "\r\n");
            richboxSend.SelectionStart = richboxSend.Text.Length;
            richboxSend.ScrollToCaret();

            if (richboxSend.TextLength > 15000)
            {
                richboxSend.Clear();
            }
        }

        private void OnSendLog(string msg)
        {
            try
            {
                this.Invoke(new DelegatLog(SendLog), new object[] { msg });
            }
            catch
            {

            }
        }
        private void Log(string msg)
        {
            String timeStr = DateTime.Now.ToString();
            richboxReceive.AppendText(msg + "\r\n");

            LogFile(timeStr + msg + "\r\n");

            richboxReceive.SelectionStart = richboxReceive.Text.Length;
            richboxReceive.ScrollToCaret();
            if (richboxReceive.TextLength > 15000)
            {
                richboxReceive.Clear();
            }
        }

        private void OnLog(string msg)
        {
            try
            {
                this.Invoke(new DelegatLog(Log), new object[] { msg });
            }
            catch
            {

            }
        }

        private void LogFile(string msg)
        {
            m_StreamWriter.WriteLine(msg);
            m_StreamWriter.Flush();
        }

        private void OpenCJBSerialPort()
        {
            try
            {
                if (m_CJBSerialPort != null && m_CJBIsConnect && m_CJBSerialPort.IsOpen)
                {
                    return;
                }

                if (m_CJBComPort == null || m_CJBBaudRate == 0)
                {
                    MessageBox.Show("请选择采集板COM口和波特率！");
                    m_CJBIsConnect = false;
                    return;
                }
                if (m_CJBComPort == m_GZHComPort && m_GZHIsConnect)
                {
                    MessageBox.Show("这串口已用了.请选择别的串口！");
                    m_CJBIsConnect = false;
                    return;
                }

                m_CJBSerialPort = new SerialPort(m_CJBComPort, m_CJBBaudRate, Parity.None, 8, StopBits.One);

                //m_CJBSerialPort.WriteTimeout = 500;
                //m_CJBSerialPort.ReadTimeout = 500;

                m_CJBSerialPort.Open();

                if (m_CJBSerialPort.IsOpen)
                {
                    Loger.send("打开采集板串口成功");
                    m_CJBIsConnect = true;
                    m_IsCJBContinueRead = true;

                    m_CJBReadThread = new Thread(CJBSerialPortRead);

                    CJBOpenStatusButtons();

                    m_CJBReadThread.Start();
                }
                else
                {
                    Loger.send("打开采集板串口失败");
                    m_CJBIsConnect = false;
                }
            }
            catch (Exception ex)
            {
                if (m_ForDevelopment) Loger.send(ex.ToString());
                m_CJBIsConnect = false;
                MessageBox.Show("打开采集板串口失败");
            }
        }
        private void CloseCJBSerialPort()
        {
            if (m_CJBIsConnect && m_CJBSerialPort != null && m_CJBSerialPort.IsOpen)
            {
                m_IsCJBContinueRead = false;

                try
                {
                    m_CJBSerialPort.DiscardInBuffer();
                    m_CJBSerialPort.DiscardOutBuffer();

                    m_CJBSerialPort.Close();

                    m_CJBSerialPort = null;

                    m_CJBIsConnect = false;

                    Loger.send("采集板串口成功关闭");
                }
                catch (Exception ex)
                {
                    if (m_ForDevelopment) Loger.send(ex.ToString());
                }
            }
        }

        private Boolean CheckCJBCom()
        {
            if (m_CJBIsConnect)
            {
                return true;
            }
            else
            {
                MessageBox.Show("请连接采集板COM口");
                return false;
            }
        }

        private void CJBSerialPortRead()
        {
            while (m_IsCJBContinueRead)
            {
                Thread.Sleep(1);
                try
                {
                    int nReviceBytesNum = m_CJBSerialPort.BytesToRead; ///收到的字节数。
                    byte[] ReadBuf = new byte[nReviceBytesNum]; ///定义接收字节数组

                    m_CJBSerialPort.Read(ReadBuf, 0, nReviceBytesNum); ///接收数据 

                    Boolean hasHeader = false;
                    foreach (byte temp in ReadBuf)
                    {
                        if (temp == m_HeaderByte)
                        {
                            m_CJBReceivedBytes.Clear();
                            hasHeader = true;
                        }
                        if (hasHeader || m_CJBReceivedBytes.Count > 0)
                        {
                            m_CJBReceivedBytes.Add(temp);
                        }
                        if (temp == m_FooterByte)
                        {
                            try
                            {
                                if (m_CJBReceivedBytes.Count > 3)
                                {
                                    DealWithCJBData(m_CJBReceivedBytes);
                                }

                            }
                            catch (Exception ex)
                            {
                                if (m_ForDevelopment) Loger.send(ex.ToString());
                            }

                            m_CJBReceivedBytes.Clear();
                            hasHeader = false;
                        }
                    }

                    //m_CJBSerialPort.DiscardInBuffer();  //清空接收缓冲区
                }
                catch (Exception ex)
                {
                    if (m_ForDevelopment) Loger.send(ex.Message);
                    m_CJBReceivedBytes.Clear();
                }
            }
        }

        private void Initionalization()
        {

            SetPortNames();

            // 0~8:档, 0~5:5,50,250,500,1K,1M 0~4:RA,LA,TT,RL,LL, 
            m_Measurements = Util.CreateJaggedArray<float[][][]>(-1, 9, 5, 6);
            m_PhaseAngles = Util.CreateJaggedArray<float[][][]>(-1, 9, 7, 3);
            m_Errors = Util.CreateJaggedArray<float[][][]>(-1, 9, 5, 6);

            IniValue2Params();

            InitButtons();

            richboxReceive.Clear();
            richboxSend.Clear();

            CloseCJBSerialPort();
            CJBOpenStatusButtons();

            m_SelectedTestWeight = -1;

            m_export = new Export();
            m_export.m_BodyStr = m_BodyStr;
            m_export.m_RateStr = m_RateStr;
            m_export.m_TestWeightLevels = m_TestWeightLevels;
            m_export.m_WeightPositionLevels = m_WeightPositionLevels;
            m_export.m_OriginValues = m_OriginValues;
            m_export.m_BodyAngleStr = m_BodyAngleStr;

            m_export1 = new Export1();
            m_export1.m_BodyStr = m_BodyStr;
            m_export1.m_RateStr = m_RateStr;
            m_export1.m_TestWeightLevels = m_TestWeightLevels;
            m_export1.m_WeightPositionLevels = m_WeightPositionLevels;
            m_export1.m_OriginValues = m_OriginValues;
            m_export1.m_BodyAngleStr = m_BodyAngleStr;
            cbxSelectGZHPosition.SelectedIndex = 0;
            cboSelectTestItemsSD.SelectedIndex = 0;
        }
        private void InitButtons()
        {
            btnCJBOpen.BackColor = Color.Red;
            btnCalibration0.BackColor = Color.LightGray;
            btnCalibration25.BackColor = Color.LightGray;
            btnCalibration50.BackColor = Color.LightGray;
            btnCalibration75.BackColor = Color.LightGray;
            btnCalibration100.BackColor = Color.LightGray;

            btnCalib1.BackColor = Color.LightGray;
            btnCalib2.BackColor = Color.LightGray;
            btnCalib3.BackColor = Color.LightGray;
            btnCalib4.BackColor = Color.LightGray;
            btnCalib5.BackColor = Color.LightGray;
            btnCalib6.BackColor = Color.LightGray;

            txtWeight.Text = "";
            btnGetWeight.BackColor = Color.LightGray;
            CalibrateBtnsEnable(false);

            btnExport.Enabled = false;
            btnExport1.Enabled = false;

            btnSDTest.BackColor = Color.LightGray;
            btnSDTest.Enabled = false;

            btnBodyTest.BackColor = Color.LightGray;
            btnBodyTest.Enabled = false;
        }

        private void CJBOpenStatusButtons()
        {
            btnCJBOpen.BackColor = m_CJBIsConnect ? Color.Green : Color.Red;

            if (m_CJBIsConnect)
            {
                CalibrateBtnsEnable(true);
                Util.cbxValue2Ini("CJB_Com", cbxCom2);
            }
        }

        private void CalibrateBtnsEnable(Boolean enable)
        {
            if (enable)
            {
                if (m_CJBIsConnect)
                {
                    btnCalibration100.Enabled = true;
                    btnCalibration75.Enabled = true;
                    btnCalibration50.Enabled = true;
                    btnCalibration25.Enabled = true;
                    btnCalibration0.Enabled = true;
                    btnGetWeight.Enabled = true;
                    btnBodyTest.Enabled = true;
                    btnSDTest.Enabled = true;
                    btnCalib1.Enabled = true;
                    btnCalib2.Enabled = true;
                    btnCalib3.Enabled = true;
                    btnCalib4.Enabled = true;
                    btnCalib5.Enabled = true;
                    btnCalib6.Enabled = true;
                }
            }
            else
            {
                btnCalibration100.Enabled = enable;
                btnCalibration75.Enabled = enable;
                btnCalibration50.Enabled = enable;
                btnCalibration25.Enabled = enable;
                btnCalibration0.Enabled = enable;
                btnGetWeight.Enabled = enable;
                btnSDTest.Enabled = enable;
                btnBodyTest.Enabled = enable;
                btnCalib1.Enabled = enable;
                btnCalib2.Enabled = enable;
                btnCalib3.Enabled = enable;
                btnCalib4.Enabled = enable;
                btnCalib5.Enabled = enable;
                btnCalib6.Enabled = enable;

            }
        }

        private void Params2Ini()
        {
            Util.cbxValue2Ini("CJB_Com", cbxCom2);
            Util.cbxValue2Ini("CJB_Baud", cboxBaud2);

            try
            {
                for (int i = 0; i < 9; i++)
                {
                    INI.WriteIni("level" + (i + 1), string.Join(",", m_OriginValues[i]));
                }
            }
            catch (Exception ex)
            {
                if (m_ForDevelopment) Loger.send(ex.ToString());
            }

            INI.WriteIni("ErrPerM1", m_ErrPerM1.ToString());
            INI.WriteIni("ErrPerM2", m_ErrPerM2.ToString());

            INI.WriteIni("TestWeightLevels", string.Join(",", m_TestWeightLevels));
            INI.WriteIni("MachineTypeList", string.Join(",", m_MachineTypeList));
        }

        private void IniValue2Params()
        {
            //COM 
            Util.IniValue2Combox("CJB_Com", cbxCom2);
            Util.IniValue2Combox("CJB_Baud", cboxBaud2);

            //电阻
            m_OriginValues = new float[9][] {
                new float[5] { 510, 750, 50, 510, 750},
                new float[5] { 360, 680, 43, 360, 680},
                new float[5] { 300, 560, 36, 300, 560},
                new float[5] { 250, 430, 30, 250, 430},
                new float[5] { 200, 300, 20, 200, 300},
                new float[5] { 150, 200, 10, 150, 200},
                new float[5] { 200, 360, 15, 200, 360},
                new float[5] { 340, 470, 27, 340, 470},
                new float[5] { 470, 620, 36, 470, 620}
            };
            for (int i = 0; i < 9; i++)
            {
                try
                {
                    string level = INI.GetValue("level" + (i + 1));
                    string[] strValues = level.Split(',');
                    if (strValues.Length == 5)
                    {
                        m_OriginValues[i] = Array.ConvertAll(strValues, float.Parse);
                    }
                }
                catch (Exception ex)
                {
                    if (m_ForDevelopment) Loger.send(ex.ToString());
                }
            }

            //误差
            try
            {
                m_ErrPerM1 = float.Parse(INI.GetValue("ErrPerM1"));
                m_ErrPerM2 = float.Parse(INI.GetValue("ErrPerM2"));
            }
            catch (Exception ex)
            {
                if (m_ForDevelopment) Loger.send(ex.ToString());
            }

            //体重测量位置
            m_WeightPositionLevels = new string[] { "中心", "前", "后", "左", "右" };
            cbxWeightPosition.Items.Clear();

            foreach (string level in m_WeightPositionLevels)
            {
                cbxWeightPosition.Items.Add(level);
            }
            cbxWeightPosition.SelectedIndex = 0;

            //体重测量选项
            m_TestWeightLevels = new float[] { 10, 15, 50, 120 };

            m_TestWeightResults = Util.CreateJaggedArray<float[][]>(-1, m_TestWeightLevels.Length, m_WeightPositionLevels.Length);

            try
            {
                string levels = INI.GetValue("TestWeightLevels");
                string[] arrLevels = levels.Split(',');
                if (arrLevels.Length > 1)
                {
                    m_TestWeightLevels = Array.ConvertAll(arrLevels, float.Parse); ;
                    m_TestWeightResults = Util.CreateJaggedArray<float[][]>(-1, arrLevels.Length, m_WeightPositionLevels.Length);
                }
            }
            catch (Exception ex)
            {
                if (m_ForDevelopment) Loger.send(ex.ToString());
            }

            cbxTestWeight.Items.Clear();

            foreach (float level in m_TestWeightLevels)
            {
                cbxTestWeight.Items.Add(level);
            }

            //设备设置
            m_MachineTypeList = new string[] { "AiNST-CNDS20", "AiNST-CNDS10", "AiNST-KNDS20", "AiNST-KNDS10" };
            try
            {
                string levels = INI.GetValue("MachineTypeList");
                string[] arrLevels = levels.Split(',');
                if (arrLevels.Length > 1)
                {
                    m_MachineTypeList = arrLevels;
                }
            }
            catch (Exception ex)
            {
                if (m_ForDevelopment) Loger.send(ex.ToString());
            }

            cbxMachineType.Items.Clear();

            foreach (string type in m_MachineTypeList)
            {
                cbxMachineType.Items.Add(type);
            }
            cbxMachineType.SelectedIndex = 0;
        }

        private void SendCommand2GZH(byte[] data)
        {
            try
            {
                if (data == null || data.Length <= 0)
                {
                    Loger.send("GZH:发送数据为空");
                    return;
                }

                Loger.send("GZH:" + Util.logBytes(data));

                if (m_GZHSerialPort != null && m_GZHSerialPort.IsOpen)
                {
                    try
                    {
                        m_GZHSerialPort.DiscardOutBuffer();
                        m_GZHSerialPort.Write(data, 0, data.Length);
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                }
                else
                {
                    MessageBox.Show("工装串口未打开");
                }

            }
            catch (Exception ex)
            {
                if (m_ForDevelopment) Loger.send(ex.Message);
            }
        }
        private void SendCommand2CJB(byte[] data)
        {            
            try
            {
                if (data == null || data.Length <= 0)
                {
                    Loger.send("CJB:发送数据为空");
                    return;
                }

                Loger.send("CJB:" + Util.logBytes(data));


                if (m_CJBSerialPort != null && m_CJBSerialPort.IsOpen)
                {
                    try
                    {
                        m_CJBSerialPort.DiscardOutBuffer();
                        m_CJBSerialPort.Write(data, 0, data.Length);
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                }
                else
                {
                    MessageBox.Show("采集板串口未打开");
                }

            }
            catch (Exception ex)
            {
                if (m_ForDevelopment) Loger.send(ex.Message);
            }
        }
        
        private void DealWithCJBData(List<byte> receivedBytes)
        {
            int length = ((receivedBytes[3] - 10) << 6) + (receivedBytes[2] - 10);

            //长度验证
            if (length != receivedBytes.Count - 6)
            {
                Loger.recv("CJB:长度不一样 " + (length + 6) + ", " + receivedBytes.Count);
                return;
            }

            byte[] bytes = new byte[length + 6];
            Array.Copy(receivedBytes.ToArray(), 0, bytes, 0, length + 6);

            if (m_ForDevelopment)
            {
                // log decrypted message
                string msg = string.Empty;
                msg = Util.logBytes(bytes);
                msg += "==>" + System.Text.Encoding.Default.GetString(bytes);
                Loger.recv("CJB:" + msg);
            }

            //CRC校验
            if (Util.getCRC(bytes, 1, length + 3) != bytes[length + 4])
            {
                Loger.recv("CJB:CRC校验失败");
                //Loger.recv("CJB:重新发送指令");
                //TODO
                //SendCommand2CJB(m_ReadStep);
                //return;
            }

            //体重校准
            if (bytes[4] == 0x4c && bytes[5] == 0x43 && bytes[6] == 0x58)
            {
                switch (m_CJBCommandStep)
                {
                    case 100:
                        btnCalibration100.BackColor = Color.Green;
                        Loger.recv("CJB:体重校准100kg完成");
                        CalibrateBtnsEnable(true);
                        break;
                    case 101:
                        btnCalibration75.BackColor = Color.Green;
                        Loger.recv("CJB:体重校准75kg完成");
                        CalibrateBtnsEnable(true);
                        break;
                    case 102:
                        btnCalibration50.BackColor = Color.Green;
                        Loger.recv("CJB:体重校准50kg完成");
                        CalibrateBtnsEnable(true);
                        break;
                    case 103:
                        btnCalibration25.BackColor = Color.Green;
                        Loger.recv("CJB:体重校准25kg完成");
                        CalibrateBtnsEnable(true);
                        break;
                    case 104:
                        btnCalibration0.BackColor = Color.Green;
                        Loger.recv("CJB:体重校准0kg完成");
                        CalibrateBtnsEnable(true);
                        break;
                    default:
                        break;
                }
                return;
            }

            //体重的读取
            if (m_CJBCommandStep == 105 && bytes[4] == 0x4c && bytes[5] == 0x43 && bytes[6] == 0x4c)
            {
                byte[] bytes1 = new byte[8];
                Array.Copy(bytes, 8, bytes1, 0, 8);
                string weightStr = System.Text.Encoding.Default.GetString(bytes1);
                float weightf = Util.toFloat(Util.exchange(weightStr));

                if (m_SelectedTestWeight > -1)
                {
                    m_TestWeightResults[m_SelectedTestWeight][m_SelectedWeightPosition] = weightf;
                }

                this.Invoke(new EventHandler(delegate
                {
                    btnGetWeight.BackColor = Color.Green;
                    Loger.recv("CJB: 体重：" + weightf.ToString());
                    CalibrateBtnsEnable(true);
                    txtWeight.Text = weightf.ToString();
                }));

                m_CJBCommandStep = -1;

                return;
            }

            //校准完成
            if (bytes[4] == 0x45 && bytes[5] == 0x57 && length > 6)
            {
                Loger.recv(String.Format("CJB: 校准{0}完成", Int32.Parse(System.Text.Encoding.Default.GetString(new byte[] { bytes[8] })) + 1));
            }

            //节段检测 sp m_counter 02  43  0c  0a  53  50  46  03
            if (bytes[4] == 0x53 && bytes[5] == 0x50)
            {
                Loger.recv("阻抗： " + string.Format("{0}", m_Counter++));
            }

            //档位校准完成
            if (bytes[4] == 0x43 && bytes[5] == 0x44 && bytes[6] == 0x52)
            {
                switch (bytes[7])
                {
                    case 0x30:
                        Loger.recv("CJB:6档位校准完成");
                        btnCalib6.BackColor = Color.Green;
                        break;
                    case 0x31:
                        Loger.recv("CJB:5档位校准完成");
                        btnCalib5.BackColor = Color.Green;
                        break;
                    case 0x32:
                        Loger.recv("CJB:4档位校准完成");
                        btnCalib4.BackColor = Color.Green;
                        break;
                    case 0x33:
                        Loger.recv("CJB:3档位校准完成");
                        btnCalib3.BackColor = Color.Green;
                        break;
                    case 0x34:
                        Loger.recv("CJB:2档位校准完成");
                        btnCalib2.BackColor = Color.Green;
                        break;
                    case 0x35:
                        Loger.recv("CJB:1档位校准完成");
                        btnCalib1.BackColor = Color.Green;
                        break;
                    default:
                        break;
                }
                CalibrateBtnsEnable(true);
                return;
            }

            //测试完成
            if (bytes[4] == 0x4d)
            {
                switch (bytes[5])
                {
                    case 0x31:
                        Loger.recv(String.Format("CJB: {0}档位：M1测试完成", m_GZHCommandStep + 1));
                        SendCommand2CJB(61);
                        break;
                    case 0x32:
                        Loger.recv(String.Format("CJB: {0}档位：M2测试完成", m_GZHCommandStep + 1));
                        SendCommand2CJB(71);
                        break;
                    case 0x33:
                        Loger.recv(String.Format("CJB: {0}档位：M3测试完成", m_GZHCommandStep + 1));
                        SendCommand2CJB(61);
                        break;
                    case 0x34:
                        Loger.recv(String.Format("CJB: {0}档位：M4测试完成", m_GZHCommandStep + 1));
                        SendCommand2CJB(71);
                        break;
                    case 0x35:
                        Loger.recv(String.Format("CJB: {0}档位：M5测试完成", m_GZHCommandStep + 1));
                        SendCommand2CJB(80);
                        break;
                    case 0x36:
                        Loger.recv(String.Format("CJB: {0}档位：M6测试完成", m_GZHCommandStep + 1));
                        this.Invoke(new EventHandler(delegate { SuccessSD(); }));
                        break;
                    default:
                        break;
                }
                
                return;
            }

            int selected = m_IsAutoTest? m_cboSelectedTestItem : m_cboSelectedTestItemSD;

            //读取阻抗
            if (bytes[4] == 0x43 && bytes[5] == 0x43)
            {
                int n = 2;
                if (m_CJBCommandStep == 61)
                {
                    Loger.recv(String.Format("CJB: {0}档位：{1} 阻抗读取成功", m_GZHCommandStep + 1, selected == 0 ? "M3" : (selected == 1 ? "M1" : "M5")));
                    n = 4;
                }
                else if (m_CJBCommandStep == 71)
                {
                    Loger.recv(String.Format("CJB: {0}档位：{1} 阻抗读取成功", m_GZHCommandStep + 1, selected == 0 ? "M4" : (selected == 1 ? "M2" : "M6")));
                    n = 2;
                }

                float[] res = new float[n * 5];

                byte[] bytes1 = new byte[n * 5 * 9];

                Array.Copy(bytes, 6, bytes1, 0, n * 5 * 9);

                for (int i = 0; i < n * 5; i++)
                {
                    byte[] zukang = new byte[8];

                    Array.Copy(bytes1, i * 9, zukang, 0, 8);

                    string str = System.Text.Encoding.Default.GetString(zukang);

                    float impedencef = Util.toFloat(Util.exchange(str));

                    res[i] = impedencef;
                }
                //输出阻抗到显示
                String impedanceStr = "\r\n";
                String zukangStr = String.Format("{0, -10}  {1, -10}  {2, -10}  {3, -10}  {4, -10}  {5, -10}\r\n", "频率/类别", m_BodyStr[0], m_BodyStr[1], m_BodyStr[2], m_BodyStr[3], m_BodyStr[4]);
                impedanceStr += zukangStr;

                //阻抗数据进行交换
                for (int i = 0; i < n; i++)
                {
                    float[] orderRes = new float[] { res[i * 5], res[i * 5 + 1], res[i * 5 + 2], res[i * 5 + 3], res[i * 5 + 4] };

                    int idx = m_CJBCommandStep == 61 ? i : (i + 4);

                    m_Measurements[m_GZHCommandStep][0][idx] = orderRes[0];
                    m_Measurements[m_GZHCommandStep][1][idx] = orderRes[1];
                    m_Measurements[m_GZHCommandStep][2][idx] = orderRes[2];
                    m_Measurements[m_GZHCommandStep][3][idx] = orderRes[3];
                    m_Measurements[m_GZHCommandStep][4][idx] = orderRes[4];
                    
                    zukangStr = String.Format("{0, -10} {1, -10}  {2, -10}  {3, -10}  {4, -10}  {5, -10}\r\n", m_RateStr[(m_CJBCommandStep == 61 ? i : (i + 4))], orderRes[0], orderRes[1], orderRes[2], orderRes[3], orderRes[4]);
                    impedanceStr += zukangStr;
                }

                Loger.recv(impedanceStr);

                if (m_CJBCommandStep == 61)
                {
                    if (selected == 1)
                    {
                        SendCommand2CJB(70);
                    }
                    else
                    {
                        SendCommand2CJB(80);
                    }
                }
                else if (m_CJBCommandStep == 71)
                {
                    if (m_IsAutoTest)
                    {
                        this.Invoke(new EventHandler(delegate { Success(); }));
                    }
                    else
                    {
                        this.Invoke(new EventHandler(delegate { SuccessSD(); }));
                    }
                }
            }

            //读取相位角
            if (bytes[4] == 0x43 && bytes[5] == 0x50)
            {
                int n = 3;
                float[] res = new float[n * 7];

                byte[] bytes1 = new byte[n * 7 * 9];

                Array.Copy(bytes, 6, bytes1, 0, n * 7 * 9);

                for (int i = 0; i < n * 7; i++)
                {
                    byte[] angle = new byte[8];

                    Array.Copy(bytes1, i * 9, angle, 0, 8);

                    string str = System.Text.Encoding.Default.GetString(angle);

                    res[i] = Util.toFloat(Util.exchange(str));
                }

                Loger.recv(String.Format("CJB: {0}档位：{1} 相位角读取成功", m_GZHCommandStep + 1, "M3"));

                //输出阻抗到显示
                String angleStr = "\r\n";
                angleStr += String.Format("{0, -10}  {1, -10}  {2, -10}  {3, -10}  {4, -10}  {5, -10}  {6, -10}  {7, -10}\r\n", "频率/类别", m_BodyAngleStr[0], m_BodyAngleStr[1], m_BodyAngleStr[2], m_BodyAngleStr[3], m_BodyAngleStr[4], m_BodyAngleStr[5], m_BodyAngleStr[6]);

                //阻抗数据进行交换
                for (int i = 0; i < n; i++)
                {
                    float[] orderRes = new float[] { res[i * 7], res[i * 7 + 1], res[i * 7 + 2], res[i * 7 + 3], res[i * 7 + 4], res[i * 7 + 5], res[i * 7 + 6] };
                    
                    m_PhaseAngles[m_GZHCommandStep][0][i] = orderRes[0];
                    m_PhaseAngles[m_GZHCommandStep][1][i] = orderRes[1];
                    m_PhaseAngles[m_GZHCommandStep][2][i] = orderRes[2];
                    m_PhaseAngles[m_GZHCommandStep][3][i] = orderRes[3];
                    m_PhaseAngles[m_GZHCommandStep][4][i] = orderRes[4];
                    m_PhaseAngles[m_GZHCommandStep][5][i] = orderRes[5];
                    m_PhaseAngles[m_GZHCommandStep][6][i] = orderRes[6];

                    angleStr += String.Format("{0, -10} {1, -10}  {2, -10}  {3, -10}  {4, -10}  {5, -10}  {6, -10}  {7, -10}\r\n", m_RateStr[i], orderRes[0], orderRes[1], orderRes[2], orderRes[3], orderRes[4], orderRes[5], orderRes[6]);
                }

                Loger.recv(angleStr);

                SendCommand2CJB(70);
            }
        }
        private void StartCalibrateImpedance()
        {
            int i = m_cboSelectedTestItem == 1 ? 0 : (m_cboSelectedTestItem == 0 ? 9 : -9);

            if (m_chkIsOnlyTest)
            {
                m_ControlStep = 1;
                SendCommand2GZH(0);
            }
            else
            {
                m_ControlStep = 0;
                SendCommand2GZH(5);
            }
        }

        private void SendCommand2GZH(int step)
        {
            m_GZHCommandStep = step;
            Thread.Sleep(1000);
            SendCommand2CJB(60);
            Thread.Sleep(1000);
        }

        private void SendCommand2CJB(int step)
        {
            m_CJBCommandStep = step;
            Thread.Sleep(2000);
            int selected = m_IsAutoTest ? m_cboSelectedTestItem : m_cboSelectedTestItemSD;
            switch (step)
            {
                case 0:
                    SendCommand2CJB(m_CJBCommand.Make(new byte[] { 0x43, 0x44, 0x52, 0x35, 0x1b }));
                    Loger.send("CJB:校准1档位");
                    break;
                case 1:
                    SendCommand2CJB(m_CJBCommand.Make(new byte[] { 0x43, 0x44, 0x52, 0x34, 0x1b }));
                    Loger.send("CJB:校准2档位");
                    break;
                case 2:
                    SendCommand2CJB(m_CJBCommand.Make(new byte[] { 0x43, 0x44, 0x52, 0x33, 0x1b }));
                    Loger.send("CJB:校准3档位");
                    break;
                case 3:
                    SendCommand2CJB(m_CJBCommand.Make(new byte[] { 0x43, 0x44, 0x52, 0x32, 0x1b }));
                    Loger.send("CJB:校准4档位");
                    break;
                case 4:
                    SendCommand2CJB(m_CJBCommand.Make(new byte[] { 0x43, 0x44, 0x52, 0x31, 0x1b }));
                    Loger.send("CJB:校准5档位");
                    break;
                case 5:
                    SendCommand2CJB(m_CJBCommand.Make(new byte[] { 0x43, 0x44, 0x52, 0x30, 0x1b }));
                    Loger.send("CJB:校准6档位");
                    break;
                case 60:
                    m_Counter = 1;
                    switch (selected)                    
                    {
                        case 0:
                            SendCommand2CJB(m_CJBCommand.Make(new byte[] { 0x4d, 0x33, 0x30, 0x31, 0x31, 0x31, 0x31, 0x30, 0x1b }));
                            Loger.send("CJB:测量M3");
                            break;
                        case 1:
                            SendCommand2CJB(m_CJBCommand.Make(new byte[] { 0x4d, 0x31, 0x30, 0x31, 0x31, 0x31, 0x31, 0x30, 0x1b }));
                            Loger.send("CJB:测量M1");
                            break;
                        case 2:
                            SendCommand2CJB(m_CJBCommand.Make(new byte[] { 0x4d, 0x35, 0x30, 0x31, 0x31, 0x31, 0x31, 0x30, 0x1b }));
                            Loger.send("CJB:测量M5");
                            break;
                        default:
                            break;
                    }
                    break;
                case 61:
                    if ((m_IsAutoTest ? m_chkIsOriginalR : m_IsOrgImpSD))
                    {
                        SendCommand2CJB(m_CJBCommand.Make(new byte[] { 0x43, 0x43, 0x30, 0x32, 0x32, 0x32, 0x32, 0x30, 0x1b }));
                    }
                    else
                    {
                        SendCommand2CJB(m_CJBCommand.Make(new byte[] { 0x43, 0x43, 0x30, 0x31, 0x31, 0x31, 0x31, 0x30, 0x1b }));
                    }
                    Loger.send(string.Format("CJB:读取{0}", selected == 0 ? "M3": (selected == 1 ? "M1" : "M5")));
                    break;
                case 70:
                    m_Counter = 1;
                    switch (selected)
                    {
                        case 0:
                            SendCommand2CJB(m_CJBCommand.Make(new byte[] { 0x4d, 0x34, 0x31, 0x30, 0x30, 0x30, 0x30, 0x31, 0x1b }));
                            Loger.send("CJB:测量M4");
                            break;
                        case 1:
                            SendCommand2CJB(m_CJBCommand.Make(new byte[] { 0x4d, 0x32, 0x31, 0x30, 0x30, 0x30, 0x30, 0x31, 0x1b }));
                            Loger.send("CJB:测量M2");
                            break;
                        case 2:
                            SendCommand2CJB(m_CJBCommand.Make(new byte[] { 0x4d, 0x36, 0x31, 0x30, 0x30, 0x30, 0x30, 0x31, 0x1b }));
                            Loger.send("CJB:测量M6");
                            break;
                        default:
                            break;
                    }
                    break;
                case 71:
                    if ((m_IsAutoTest ? m_chkIsOriginalR : m_IsOrgImpSD))
                    {
                        SendCommand2CJB(m_CJBCommand.Make(new byte[] { 0x43, 0x43, 0x32, 0x30, 0x30, 0x30, 0x30, 0x32, 0x1b }));
                    }
                    else
                    {
                        SendCommand2CJB(m_CJBCommand.Make(new byte[] { 0x43, 0x43, 0x31, 0x30, 0x30, 0x30, 0x30, 0x31, 0x1b }));
                    }
                    Loger.send(string.Format("CJB:读取{0}", selected == 0 ? "M4" : (selected == 1 ? "M2" : "M6")));
                    break;
                case 80:
                    SendCommand2CJB(m_CJBCommand.Make(new byte[] { 0x43, 0x50, 0x30, 0x31, 0x31, 0x31, 0x30, 0x30, 0x1b }));
                    Loger.send("CJB:读取相位角");
                    break;
                case 100:
                    Loger.send("CJB:体重校准100kg");
                    SendCommand2CJB(m_CJBCommand.Make(new byte[] { 0x4c, 0x43, 0x43, 0x1b }));
                    Thread.Sleep(2000);
                    SendCommand2CJB(m_CJBCommand.Make(new byte[] { 0x4c, 0x43, 0x34, 0x1b }));
                    Thread.Sleep(2000);
                    SendCommand2CJB(m_CJBCommand.Make(new byte[] { 0x4c, 0x43, 0x58, 0x1b }));
                    break;
                case 101:
                    SendCommand2CJB(m_CJBCommand.Make(new byte[] { 0x4c, 0x43, 0x43, 0x1b }));
                    Thread.Sleep(2000);
                    SendCommand2CJB(m_CJBCommand.Make(new byte[] { 0x4c, 0x43, 0x33, 0x1b }));
                    Thread.Sleep(2000);
                    SendCommand2CJB(m_CJBCommand.Make(new byte[] { 0x4c, 0x43, 0x58, 0x1b }));
                    Loger.send("CJB:体重校准75kg");
                    break;
                case 102:
                    SendCommand2CJB(m_CJBCommand.Make(new byte[] { 0x4c, 0x43, 0x43, 0x1b }));
                    Thread.Sleep(2000);
                    SendCommand2CJB(m_CJBCommand.Make(new byte[] { 0x4c, 0x43, 0x32, 0x1b }));
                    Thread.Sleep(2000);
                    SendCommand2CJB(m_CJBCommand.Make(new byte[] { 0x4c, 0x43, 0x58, 0x1b }));
                    Loger.send("CJB:体重校准50kg");
                    break;
                case 103:
                    SendCommand2CJB(m_CJBCommand.Make(new byte[] { 0x4c, 0x43, 0x43, 0x1b }));
                    Thread.Sleep(2000);
                    SendCommand2CJB(m_CJBCommand.Make(new byte[] { 0x4c, 0x43, 0x31, 0x1b }));
                    Thread.Sleep(2000);
                    SendCommand2CJB(m_CJBCommand.Make(new byte[] { 0x4c, 0x43, 0x58, 0x1b }));
                    Loger.send("CJB:体重校准25kg");
                    break;
                case 104:
                    SendCommand2CJB(m_CJBCommand.Make(new byte[] { 0x4c, 0x43, 0x43, 0x1b }));
                    Thread.Sleep(2000);
                    SendCommand2CJB(m_CJBCommand.Make(new byte[] { 0x4c, 0x43, 0x30, 0x1b }));
                    Thread.Sleep(2000);
                    SendCommand2CJB(m_CJBCommand.Make(new byte[] { 0x4c, 0x43, 0x58, 0x1b }));
                    Loger.send("CJB:体重校准0kg");
                    break;
                case 105:
                    SendCommand2CJB(m_CJBCommand.Make(new byte[] { 0x4c, 0x43, 0x4c, 0x1b }));
                    Loger.send("CJB:读取体重");
                    break;
                default:
                    break;
            }
        }
        private void Success()
        {
            btnExport.Enabled = true;
            btnExport1.Enabled = true;

            CalibrateBtnsEnable(true);

            Boolean checkResult = true;
            rtbResult.Clear();
            
            for (int i = 0; i < 9; i++)
            {
                if (!CheckOneLevel(i))
                {
                    checkResult = false;
                }
            }
                
            if (checkResult)
            {
                Loger.recv("完成");
                MessageBox.Show("完成");
            }
            else
            {
                Loger.recv("完成，测试结果不好");
                MessageBox.Show("完成，测试结果不好");
            }
        }

        private void SetPortNames()
        {
            string[] ports = SerialPort.GetPortNames();

            cbxCom2.Items.Clear();

            foreach (string port in ports)
            {
                cbxCom2.Items.Add(port);
            }
        }

        private string GetMachineTypePre()
        {
            string[] tmp = m_MachineType.Split('-');
            string type = "CN";
            if (tmp.Length > 1)
            {
                try
                {
                    type = tmp[1].Substring(0, 2);
                }
                catch (Exception ex)
                {
                }
            }
            return type;
        }

        private void AutoCreateSerialNumber()
        {
            DateTime nowDate = DateTime.Now;
            string yyMMdd = nowDate.ToString("yyMMdd");
            dupSerialNumber.Items.Clear();
            string type = GetMachineTypePre();
            string selected = "";
            for (int i = 0; i < 30; i++)
            {
                string tail = (i + 1).ToString();
                string item = string.Format("{0}20{1}{2}", type, yyMMdd, tail.PadLeft(3, '0'));
                dupSerialNumber.Items.Add(item);
                if (i == 0)
                {
                    selected = item;
                }
            }
            dupSerialNumber.Text = selected;
        }

        private void StartSDTest()
        {
            m_IsBodyTest = false;
            m_IsOnlyTestSD = true;
            
            m_ControlStep = 1;
            SendCommand2GZH(m_cbxSelectGZHPosition);
        }

        private void StartBodyTest()
        {
            m_IsBodyTest = true;
            m_GZHCommandStep = 0;
            SendCommand2CJB(60);
        }

        private void SuccessSD()
        {
            btnExport.Enabled = true;
            btnExport1.Enabled = true;

            CheckOneLevel(m_GZHCommandStep);

            CalibrateBtnsEnable(true);

            if (m_IsBodyTest)
            {
                btnBodyTest.BackColor = Color.Green;
            }
            else
            {
                btnSDTest.BackColor = Color.Green;
            }
            Loger.recv("完成");
            MessageBox.Show("完成");
        }

        private Boolean CheckOneLevel(int level)
        {
            int i = level;
            Boolean isOk = true;
            int selected = m_IsAutoTest ? m_cboSelectedTestItem : m_cboSelectedTestItemSD;
            String resultStr = "";
            String errorStr = "";

            if (selected != 2)
            {
                for (int j = 0; j < 5; j++)
                {
                    for (int k = 0; k < 6; k++)
                    {
                        m_Errors[i][j][k] = (m_Measurements[i][j][k] - m_OriginValues[i][j]) / m_OriginValues[i][j] * 100;
                        float err = i < 6 ? m_ErrPerM1 : m_ErrPerM2;
                        if( j == 2)
                        {
                            err = m_ErrPerM2;
                        }
                        if (Math.Abs(m_Errors[i][j][k]) > err)
                        {
                            isOk = false;
                        }
                    }
                }

                rtbResult.AppendText("---------------------" + (i + 1) + "档位测量---------------------\r\n");                
                String headerStr = String.Format("{0, -10}  {1, -10}  {2, -10}  {3, -10}  {4, -10}  {5, -10}\r\n", "频率/类别", m_BodyStr[0], m_BodyStr[1], m_BodyStr[2], m_BodyStr[3], m_BodyStr[4]);
                resultStr += headerStr;
                errorStr += headerStr;
            }
            
            String angleStr = String.Format("{0, -10}  {1, -10}  {2, -10}  {3, -10}  {4, -10}  {5, -10}  {6, -10}  {7, -10}\r\n", "频率/类别", m_BodyAngleStr[0], m_BodyAngleStr[1], m_BodyAngleStr[2], m_BodyAngleStr[3], m_BodyAngleStr[4], m_BodyAngleStr[5], m_BodyAngleStr[6]);
            for (int j = 0; j < 6; j++)
            {
                if (selected != 2)
                {
                    resultStr += String.Format("{0, -10} {1, -10}  {2, -10}  {3, -10}  {4, -10}  {5, -10}\r\n", m_RateStr[j], m_Measurements[i][0][j], m_Measurements[i][1][j], m_Measurements[i][2][j], m_Measurements[i][3][j], m_Measurements[i][4][j]);
                    errorStr += String.Format("{0, -10} {1, -10}  {2, -10}  {3, -10}  {4, -10}  {5, -10}\r\n", m_RateStr[j], m_Errors[i][0][j], m_Errors[i][1][j], m_Errors[i][2][j], m_Errors[i][3][j], m_Errors[i][4][j]);
                }
                    
                if (selected != 1 && j < 3)
                {
                    angleStr += String.Format("{0, -10} {1, -10}  {2, -10}  {3, -10}  {4, -10}  {5, -10}  {6, -10}  {7, -10}\r\n", m_RateStr[j], m_PhaseAngles[i][0][j], m_PhaseAngles[i][1][j], m_PhaseAngles[i][2][j], m_PhaseAngles[i][3][j], m_PhaseAngles[i][4][j], m_PhaseAngles[i][5][j], m_PhaseAngles[i][6][j]);
                }
            }
            if (selected != 2)
            {
                rtbResult.AppendText(resultStr);

                if (!m_IsBodyTest)
                {
                    rtbResult.AppendText("---------------------" + (i + 1) + "档位误差---------------------\r\n");
                    if (!isOk)
                    {
                        rtbResult.SelectionStart = rtbResult.TextLength;
                        rtbResult.SelectionLength = 0;
                        rtbResult.SelectionColor = Color.Red;
                        rtbResult.AppendText(errorStr);
                        rtbResult.SelectionColor = rtbResult.ForeColor;
                    }
                    else
                    {
                        rtbResult.AppendText(errorStr);
                    }
                }
            }

            if (selected != 1)
            {
                rtbResult.AppendText("---------------------" + (i + 1) + "档位相位角---------------------\r\n");
                rtbResult.AppendText(angleStr);
            }
            return isOk;
        }


        private void cboSelectTestItemsSD_SelectedIndexChanged(object sender, EventArgs e)
        {
            m_cboSelectedTestItemSD = cboSelectTestItemsSD.SelectedIndex;
        }

        private void btnCalib1_Click(object sender, EventArgs e)
        {
            btnCalib1.BackColor = Color.Red;
            CalibrateBtnsEnable(false);
            SendCommand2CJB(0);
        }

        private void btnCalib2_Click(object sender, EventArgs e)
        {
            btnCalib2.BackColor = Color.Red;
            CalibrateBtnsEnable(false);
            SendCommand2CJB(1);
        }

        private void btnCalib3_Click(object sender, EventArgs e)
        {
            btnCalib3.BackColor = Color.Red;
            CalibrateBtnsEnable(false);
            SendCommand2CJB(2);
        }

        private void btnCalib4_Click(object sender, EventArgs e)
        {
            btnCalib4.BackColor = Color.Red;
            CalibrateBtnsEnable(false);
            SendCommand2CJB(3);
        }

        private void btnCalib5_Click(object sender, EventArgs e)
        {
            btnCalib5.BackColor = Color.Red;
            CalibrateBtnsEnable(false);
            SendCommand2CJB(4);
        }

        private void btnCalib6_Click(object sender, EventArgs e)
        {
            btnCalib6.BackColor = Color.Red;
            CalibrateBtnsEnable(false);
            SendCommand2CJB(5);
        }

        
    }
}
