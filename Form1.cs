using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO.Ports;
using Common;
using Common.Do;
using System.IO;
using System.Threading;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;
using DevComponents.DotNetBar;

namespace comTest
{
    //public delegate void DelegatLog(string msg);
    public delegate void DelegatBack(Back back);
    public partial class Form1 : RibbonForm
    {
        private string m_comPort;
        private byte m_baud;
        private byte m_baud08;
        private SerialPort m_serialPort=null;
        private byte m_gender;
        List<byte> m_recvBytes=new List<byte>();
        FileStream fs = null;
        StreamWriter sw = null;
        private bool m_isWorkStyle = true;
        private byte[] m_bytes = null;
        private int m_calibariton_type = 0;
        private bool m_isAscii = true;

        private float m_LAStand, m_RAStand, m_TTStand, m_LLStand, m_RLStand;
        private float m_LAP, m_RAP, m_TTP, m_LLP, m_RLP;
        private bool m_forDevelopment=false;
        private int m_ReadStep = 0;

        //记录阻抗数量
        private int m_counter;

        public Form1()
        {
            InitializeComponent();
            Loger.LogRecived += new DelegateLog(OnLog);
            Loger.LogSend += new DelegateLog(OnSendLog);
            try
            {
                fs = new FileStream("log.txt", FileMode.OpenOrCreate);
                sw = new StreamWriter(fs, Encoding.Unicode);
            }
            catch (Exception ex)
            {
                Loger.send(ex.ToString());
            }

        }
        private void sendLog(string msg)
        {
            richBoxSend.Text +=msg + "\r\n";
            richBoxSend.SelectionStart = richBoxSend.Text.Length;
            richBoxSend.ScrollToCaret();

            if (richBoxSend.TextLength >5000) richBoxSend.Clear();
        }

        private void OnSendLog(string msg)
        {
            try
            {
                this.Invoke(new DelegatLog(sendLog), new object[] { msg });
            }
            catch
            {

            }
        }
        private void Log(string msg)
        {
            String timeStr=DateTime.Now.ToString();
            richboxReceive.Text +=msg + "\r\n";

            logFile(timeStr + msg + "\r\n");

            richboxReceive.SelectionStart = richboxReceive.Text.Length;
            richboxReceive.ScrollToCaret();
            if (richboxReceive.TextLength >5000) richboxReceive.Clear();
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
        private void cbxCom_TabIndexChanged(object sender, EventArgs e)
        {
            m_comPort = cbxCom.SelectedText;
        }

        private void cbxBaud_SelectedIndexChanged(object sender, EventArgs e)
        {
            //1khz
            //5khz
            //10khz
            //20khz
            //50khz
            //100khz
            //200khz
            //250khz
            //500khz
            //1000khz
   
        }
        /// <summary>
        /// 发送命令
        /// </summary>
        /// <param name="SendData">发送数据</param>
        /// <param name="ReceiveData">接收数据</param>

        /// <param name="Overtime">超时时间</param>
        /// <returns></returns>
        public int SendCommand(byte[] SendData)
        {
            if (m_serialPort != null && m_serialPort.IsOpen)
            {
                try
                {
                    m_serialPort.WriteTimeout = 500;
                    m_serialPort.ReadTimeout = 500;

                    m_serialPort.DiscardInBuffer();  //清空接收缓冲区               
                    m_serialPort.Write(SendData, 0, SendData.Length);

                    return 1;
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
            else
            {
                MessageBox.Show("串口未打开");
            }
            return -1;
        }
        //异或校验码生成
        private byte getCRC(byte [] value, int index, int number)
        {
            int res = 0;
            for (int i = 0; i < number; i++)
            {
                res += value[i + index];
            }
            return (byte)((res & 0x3F) + 10);
        }

        private void btnSender_Click(object sender, EventArgs e)
        {
            try
            {
                if (m_serialPort!=null && m_serialPort.IsOpen)
                {
                    byte[] sendData = new byte[6];
                    sendData[0] = 0x64;
                    sendData[1] = 0x95;
                    sendData[2] = 0x02;
                    sendData[3] = 0x04;
                    sendData[4] = m_baud;
                    sendData[5] = getCRC(sendData, 0, sendData.Length-1);

                    string msg = string.Empty;
                    foreach (byte temp in sendData)
                    {
                        msg += temp.ToString("x2");
                        msg += "  ";
                    }
                    Loger.send(string.Format("发送：{0}", msg));

                    int n = SendCommand(sendData);
                }
                else
                    Loger.send("串口未打开");
            }
            catch (Exception ex)
            {
                Loger.send("打开串口失败！");
                Loger.send(ex.Message);
            }
        }

        private void cbxCom_SelectedValueChanged(object sender, EventArgs e)
        {
            m_comPort = cbxCom.Text;
        }

        private void cbxBaud_SelectedValueChanged(object sender, EventArgs e)
        {
            //1khz
            //5khz
            //10khz
            //20khz
            //50khz
            //100khz
            //200khz
            //250khz
            //500khz
            //1000khz
            //switch (cbxBaud.Text.Trim())
            //{
            //    case "1khz":
            //        m_baud = 0;
            //        break;
            //    case "5khz":
            //        m_baud = 1;
            //        break;
            //    case "10khz":
            //        m_baud = 2;
            //        break;
            //    case "20khz":
            //        m_baud = 3;
            //        break;
            //    case "50khz":
            //        m_baud = 4;
            //        break;
            //    case "100khz":
            //        m_baud = 5;
            //        break;
            //    case "200khz":
            //        m_baud = 6;
            //        break;
            //    case "250khz":
            //        m_baud = 7;
            //        break;
            //    case "500khz":
            //        m_baud = 8;
            //        break;
            //    case "1000khz":
            //        m_baud = 9;
            //        break;
            //}
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            cbxCom.SelectedIndex = 0;
            ////cbxBaud.SelectedIndex = 0;
            //cbxGender.SelectedIndex = 0;
            ////cbxBaud08.SelectedIndex = 0;
            //cbxWork.SelectedIndex = 0;
            cboxBaud.SelectedIndex = 2;
            //radiobtnAscii.Checked = true;

            IniValue2Text("LA", txtLAIm);
            IniValue2Text("RA", txtRAIm);
            IniValue2Text("TT", txtTTIm);
            IniValue2Text("LL", txtLLIm);
            IniValue2Text("RL", txtRLIm);

            IniValue2Text("LAP", txtLAP);
            IniValue2Text("RAP", txtRAP);
            IniValue2Text("TTP", txtTTP);
            IniValue2Text("LLP", txtLLP);
            IniValue2Text("RLP", txtRLP);

            IniValue2Combox("com", cbxCom);
            IniValue2Combox("baud", cboxBaud);

        }
        private void IniValue2Combox(string key, ComboBox cbx)
        {
            try
            {
                cbx.Text = INI.GetValue(key);
            }
            catch (Exception ex)
            {
                Loger.send(ex.ToString());
            }
        }

        private void IniValue2Text(string key, TextBox txtb)
        {
            try
            {
                txtb.Text = INI.GetValue(key);
            }
            catch (Exception ex)
            {
                Loger.send(ex.ToString());
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                if (m_serialPort != null)
                {
                    m_serialPort.DataReceived -= new SerialDataReceivedEventHandler(m_serialPort_DataReceived);
                    m_serialPort.Close();
                }
                int baudRate = 0;
                try
                {
                    baudRate = Int32.Parse(cboxBaud.Text);
                    m_comPort = cbxCom.Text;
                }
                catch (Exception ex)
                {
                    Loger.send(ex.ToString());
                    return;
                }
                m_serialPort = new SerialPort(m_comPort, baudRate, Parity.None, 8, StopBits.One);
                m_serialPort.Open();
                
                m_serialPort.DataReceived+=new SerialDataReceivedEventHandler(m_serialPort_DataReceived);

                if (m_serialPort.IsOpen)
                {
                    Loger.send("打开串口成功");
                }
                else
                    Loger.send("打开串口失败");

            }
            catch(Exception ex)
            {
                Loger.send(ex.ToString());
            }
        }
        private void m_serialPort_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            try
            {
                Thread.Sleep(300);
  
                int nReviceBytesNum = m_serialPort.BytesToRead; ///收到的字节数。
                byte[] ReadBuf = new byte[nReviceBytesNum]; ///定义接收字节数组
                m_serialPort.Read(ReadBuf, 0, nReviceBytesNum); ///接收数据 

                foreach (byte temp in ReadBuf) m_recvBytes.Add(temp);
                if (m_recvBytes.Count <= 0) return;

                if (m_recvBytes.Count > 500) m_recvBytes.Clear();

                List<int> indexList = new List<int>();
                //查找0xaa
                int index = -1;
                foreach (byte item in m_recvBytes)
                {
                    index++;
                    if (item == (byte)0x02/* && (m_recvBytes[index + 1] == 0x53 || m_recvBytes[index + 1] == 0x43)*/)
                    {
                        indexList.Add(index);
                    }
                }
                foreach (int item in indexList)
                {
                    try
                    {
                        dealWithData(item);
                    }
                    catch (Exception ex)
                    {
                        Loger.send(ex.ToString());
                    }

                }
                m_recvBytes.Clear();
            }
            catch (Exception ex)
            {
                Loger.recv(ex.Message);
            }
        }

        private float toFloat(string str)
        {
            uint num = uint.Parse(str, System.Globalization.NumberStyles.AllowHexSpecifier);
            byte[] floatVals = BitConverter.GetBytes(num);
            float f = BitConverter.ToSingle(floatVals, 0);
            return f;
        }
        //交换字符串
        private string exchange(string temp)
        {
            string value = "";

            char[] cs = temp.ToCharArray();

            //交换
            char cTemp;
            cTemp = cs[0];
            cs[0] = cs[6];
            cs[6] = cTemp;

            cTemp = cs[1];
            cs[1] = cs[7];
            cs[7] = cTemp;

            cTemp = cs[2];
            cs[2] = cs[4];
            cs[4] = cTemp;

            cTemp = cs[3];
            cs[3] = cs[5];
            cs[5] = cTemp;

            value = new string(cs);
            return value;
        }
        private void dealWithData(int index1)
        {
            int length = ((m_recvBytes[index1 + 3] - 10) << 6) + (m_recvBytes[index1 + 2] - 10);
            if (length > m_recvBytes.Count - 6)
            {
                return;
            }
            byte[] bytes = new byte[length+6];
            Array.Copy(m_recvBytes.ToArray(), index1, bytes,0, length+6);

            //输出到日志
            string logMsg = String.Empty;

            // Decrypt message
            //codec(bytes, 0);

            // log decrypted message
            string msg = string.Empty ;

            msg = logBytes(bytes);

            msg += "==>" + System.Text.Encoding.Default.GetString(bytes);
            index1 = 0;

            //lcn
            //if (bytes[4] == 0x4c && bytes[ 5] == 0x43 && bytes[ 6] == 0x4E)
            //{
            //    int a = 5;
            //}

            if(m_forDevelopment)Loger.recv(msg);

                //lcx 02  d4  0e  0a     4c 43  58      1b     32  03
           
            if (bytes[index1 + 4] == 0x4c && bytes[index1 + 5] == 0x43 && bytes[index1 + 6] == 0x58)
                {
                    switch (m_calibariton_type)
                    {
                        case 0:
                           btnCalibration0.TextColor = Color.Green;
                           Loger.recv("体重校准0kg完成");
                           break;
                        case 25:
                           btnCalibration0.TextColor = Color.Green;
                           Loger.recv("体重校准25kg完成");
                           break;
                        case 50:
                           btnCalibration0.TextColor = Color.Green;
                           Loger.recv("体重校准50kg完成");
                           break;
                        case 75:
                           btnCalibration0.TextColor = Color.Green;
                           Loger.recv("体重校准75kg完成");
                           break;
                        case 100:
                            btnCalibration0.TextColor = Color.Green;
                            Loger.recv("体重校准100kg完成");
                            break;
                        default:
                            break;
                    }
                             
                }

            //EW 45  57
            if (bytes[index1 + 4] == 0x45 && bytes[index1 + 5] == 0x57)
            {
                if (!m_forDevelopment)
                    Loger.recv("校准： " + string.Format("{0}", m_counter++));
            }

            //节段检测 sp m_counter 02  43  0c  0a  53  50  46  03
            if (bytes[index1 + 4] == 0x53 && bytes[index1 + 5] == 0x50)
            {
                if (!m_forDevelopment)
                Loger.recv("阻抗： " + string.Format("{0}", m_counter++));
            }

            //体重的读取
            if (bytes[index1 + 4] == 0x4c && bytes[index1 + 5] == 0x43 && bytes[index1 + 6] == 0x4c)
            {
                byte[] bytes1 = new byte[8];
                Array.Copy(bytes, index1 + 8, bytes1, 0, 8);
                string weightStr = System.Text.Encoding.Default.GetString(bytes1);
                float weightf = toFloat(exchange(weightStr));

                this.Invoke(new EventHandler(delegate
                {
                    btnWeight.TextColor = Color.Green;
                    Loger.recv("体重：" + weightf.ToString());
                    textWeight.Text = weightf.ToString();
                }));
            }
            //位相角的读取
            if (bytes[index1 + 4] == 0x56 && bytes[index1 + 5] == 0x48)
            {
                byte[] bytes1 = new byte[9];
                Array.Copy(bytes, index1 + 6, bytes1, 0, 8);
                string Str = System.Text.Encoding.Default.GetString(bytes1);
                string phaseStr = exchange(Str);
                float phasef = toFloat(phaseStr);
                
                this.Invoke(new EventHandler(delegate
                {
                    //btnWeight.TextColor = Color.Green;
                    Loger.recv("位相角：" + phasef.ToString());
                    PHA_TT.Text = phasef.ToString();
                }));
            }

            if (bytes[index1 + 4] == 0x56 && bytes[index1 + 5] == 0x49)
            {
                byte[] bytes1 = new byte[9];
                Array.Copy(bytes, index1 + 6, bytes1, 0, 8);
                string Str = System.Text.Encoding.Default.GetString(bytes1);
                string phaseStr = exchange(Str);
                float phasef = toFloat(phaseStr);

                this.Invoke(new EventHandler(delegate
                {
                    //btnWeight.TextColor = Color.Green;
                    Loger.recv("位相角：" + phasef.ToString());
                    PHA_TT.Text = phasef.ToString();
                }));
            }
            //测量完成M1 02  43  0c  0a  4d  31  21  03 
            if (bytes[index1 + 4] == 0x4d && bytes[index1 + 5] == 0x31)
            {
                Loger.recv("M1阻抗检测完成");

                btnM1.TextColor = Color.Green;
                clearImpedance();
                send(new byte[] { 0x02, 0x53, 0x13, 0x0a, 0x43, 0x43, 0x30, 0x31, 0x31, 0x31, 0x31, 0x30, 0x1b, 0x3f, 0x03 }); // CC
            }
            //测量完成M2 02  43  0c  0a  4d  32  21  03 
            if (bytes[index1 + 4] == 0x4d && bytes[index1 + 5] == 0x32)
            {
                Loger.recv("M2 impedance m complate");

                send(new byte[] { 0x02, 0x53, 0x13, 0x0a, 0x43, 0x43, 0x31, 0x30, 0x30, 0x30, 0x30, 0x31, 0x1b, 0x3d, 0x03 }); // CC1
            }
            //cdr0  02  c0  13  0a  43  44  52  30  31  31  31  30  1b  21  03
            if (bytes[index1 + 4] == 0x43 && bytes[index1 + 5] == 0x44 && bytes[index1 + 6] == 0x52 && bytes[index1 + 7] == 0x30)
            {
                Loger.recv("档位3校准完成");
                btnCDR0.TextColor = Color.Green;
            }
            //cdr0  02  c0  13  0a  43  44  52  30  31  31  31  30  1b  21  03
            if (bytes[index1 + 4] == 0x43 && bytes[index1 + 5] == 0x44 && bytes[index1 + 6] == 0x52 && bytes[index1 + 7] == 0x31)
            {
                Loger.recv("档位2校准完成");
                btnCDR1.TextColor = Color.Green;
            }
            //cdr0  02  c0  13  0a  43  44  52  30  31  31  31  30  1b  21  03
            if (bytes[index1 + 4] == 0x43 && bytes[index1 + 5] == 0x44 && bytes[index1 + 6] == 0x52 && bytes[index1 + 7] == 0x32)
            {
                Loger.recv("档位1校准完成");
                btnCDR2.TextColor = Color.Green;
            }

            //读取阻抗
            if (bytes[index1 + 4] == 0x43 && bytes[index1 + 5] == 0x43) //CC command
            {
                int n = 2;
                if (m_ReadStep == 1 || m_ReadStep == 3)
                {
                    Loger.recv("M1 阻抗读取成功");
                    n = 4;
                }
                else if(m_ReadStep == 2)
                {
                    Loger.recv("M2 阻抗读取成功");
                    n = 2;
                }
                float[] res=new float[n * 5];

                byte[]bytes1 = new byte[n * 5 * 9];
                Array.Copy(bytes, index1 + 6, bytes1, 0, n * 5 * 9);
                for (int i = 0; i < n * 5; i++)
                {
                    byte[] zukang = new byte[8];
                    Array.Copy(bytes1, i * 9, zukang, 0, 8);
                    string str = System.Text.Encoding.Default.GetString(zukang);
                    string impedenceStr = exchange(str);
                    float impedencef = toFloat(impedenceStr);

                    res[i] = impedencef;
                }

                //阻抗数据进行交换
                for (int i = 0; i < n; i++)
                {
                    float temp = res[i * 5];
                    temp = res[i * 5];
                    res[i * 5] = res[i * 5+4];
                    res[i * 5 + 4] = temp;
                }
                //阻抗数据进行交换
                for (int i = 0; i < n; i++)
                {
                    float temp = res[i * 5+3];
                    res[i * 5+3] = res[i * 5 + 4];
                    res[i * 5 + 4] = temp;
                }
                //阻抗数据进行交换
                for (int i = 0; i < n; i++)
                {
                    float temp = res[i * 5];
                    temp = res[i * 5];
                    res[i * 5] = res[i * 5 + 1];
                    res[i * 5 + 1] = temp;
                }
                //输出阻抗到显示
                String impedanceStr = DateTime.Now.ToString()+ "-------------------------------\r\n";
                int j=0;
                foreach (float item in res)
                {

                    if (j % 5 == 0)
                    {
                        String zukangStr = String.Format("{0},  {1},  {2} , {3}, {4}\r\n", res[j + 0], res[j + 1], res[j + 2], res[j + 3], res[j + 4]);
                        impedanceStr += zukangStr;       
                    }
                    j++;
                }

                if (m_ReadStep == 1 || m_ReadStep == 3)
                {
                    this.Invoke(new EventHandler(delegate
                    {
                    //richboxImpedance.Text += impedanceStr;
                    //if (richboxImpedance.Text.Length > 5000) richboxImpedance.Clear();
                    //richboxImpedance.SelectionStart = richboxImpedance.Text.Length;
                    //richboxImpedance.ScrollToCaret();
                    labelRA5.Text = res[0].ToString();
                        labelRA50.Text = res[5].ToString();
                        labelRA250.Text = res[10].ToString();
                        labelRA500.Text = res[15].ToString();

                        labelLA5.Text = res[1].ToString();
                        labelLA50.Text = res[6].ToString();
                        labelLA250.Text = res[11].ToString();
                        labelLA500.Text = res[16].ToString();

                        labelTT5.Text = res[2].ToString();
                        labelTT50.Text = res[7].ToString();
                        labelTT250.Text = res[12].ToString();
                        labelTT500.Text = res[17].ToString();

                        labelRL5.Text = res[3].ToString();
                        labelRL50.Text = res[8].ToString();
                        labelRL250.Text = res[13].ToString();
                        labelRL500.Text = res[18].ToString();

                        labelLL5.Text = res[4].ToString();
                        labelLL50.Text = res[9].ToString();
                        labelLL250.Text = res[14].ToString();
                        labelLL500.Text = res[19].ToString();
                                               
                        //PHA_LA.Text = res[20].ToString();
                        //PHA_RA.Text = res[21].ToString();
                        //PHA_TT.Text = res[22].ToString();
                        //PHA_LL.Text = res[23].ToString();
                        //PHA_RL.Text = res[24].ToString();
                    }));

                    if(m_ReadStep == 1)
                        // send M2
                        send(new byte[] { 0x02, 0x53, 0x13, 0x0a, 0x4d, 0x32, 0x31, 0x30, 0x30, 0x30, 0x30, 0x31, 0x1b, 0x36, 0x03 }); // M2
                    else
                        send(new byte[] { 0x02, 0x53, 0x13, 0x0a, 0x43, 0x43, 0x31, 0x30, 0x30, 0x30, 0x30, 0x31, 0x1b, 0x3d, 0x03 }); // CC1
                    m_ReadStep = 2;
                }
                else if (m_ReadStep == 2)
                {
                    this.Invoke(new EventHandler(delegate
                    {
                        btnReadImpedence.TextColor = Color.Green;
                        btnM1.TextColor = Color.Green;

                        //richboxImpedance.Text += impedanceStr;
                        //if (richboxImpedance.Text.Length > 5000) richboxImpedance.Clear();
                        //richboxImpedance.SelectionStart = richboxImpedance.Text.Length;
                        //richboxImpedance.ScrollToCaret();
                        labelRA1K.Text = res[0].ToString();
                        labelRA1M.Text = res[5].ToString();

                        labelLA1K.Text = res[1].ToString();
                        labelLA1M.Text = res[6].ToString();

                        labelTT1K.Text = res[2].ToString();
                        labelTT1M.Text = res[7].ToString();

                        labelRL1K.Text = res[3].ToString();
                        labelRL1M.Text = res[8].ToString();

                        labelLL1K.Text = res[4].ToString();
                        labelLL1M.Text = res[9].ToString();
                    }));
                    m_ReadStep = 0;
                }
            }
        }

        private byte getrandom(byte seed)
        {
            return (byte)(((((UInt32)seed * 31421 + 6927)) & 0x7F) + 0x80);
        }

        private void codec(byte[] recvBytes, int index = 0)
        {
            byte key = recvBytes[index + 1];
            int length = ((recvBytes[index + 3] - 10) << 6) + (recvBytes[index + 2] - 10);
            for (int pos = 0; pos < length; pos++)
            {
                key = getrandom(key);
                byte codecbyte = (byte)(recvBytes[index + pos + 4] ^ key);
                recvBytes[index + pos + 4] = codecbyte;
            }
        }

        private void funCode4(List<byte> recvBytes, int index)
        {
            Back4 back = new Back4();
            if (m_recvBytes[index + back.Length + 3] == getCRC(recvBytes.ToArray(), index, back.Length + 3))
            {
                //判断不同的关键词，调用不同类的处理方法

                byte[] datas = new byte[back.Length];

                Array.Copy(m_recvBytes.ToArray(), index + 3, datas, 0, back.Length);

                //根据不同的功能码，获取不同的返回数据
                back.DataBytes = datas;

                string msg = string.Format("{0}, 频段号,{1},体重,{2},阻抗1,{3},阻抗2,{4},阻抗3,{5},阻抗4,{6},阻抗5,{7},阻抗6,{8}",
                                            DateTime.Now.ToString("HH:mm:ss.fff"), back.频段号, back.体重, back.阻抗1, back.阻抗2, back.阻抗3, back.阻抗4, back.阻抗5, back.阻抗6);
                logFile(msg);
                Loger.recv(string.Format("接受：{0}", msg));

                this.Invoke(new DelegatBack(backRecieved), new object[] { back });
                m_recvBytes.RemoveRange(0, index + back.Length + 4);
            }
            //删除前21个字节
            m_recvBytes.RemoveRange(0, index + 1);
        }

        private void funCode2(List<byte> recvBytes, int index)
        {
            Back2 back = new Back2();
            if (m_recvBytes[index + back.Length + 3] == getCRC(recvBytes.ToArray(), index, back.Length + 3))
            {
                //判断不同的关键词，调用不同类的处理方法

                byte[] datas = new byte[back.Length];
                Array.Copy(m_recvBytes.ToArray(), index + 3, datas, 0, back.Length);

                //根据不同的功能码，获取不同的返回数据
                back.DataBytes = datas;

                string msg = string.Format("{0}", back.codeId);
                logFile(DateTime.Now.ToString("HH:mm:ss.fff")+": "+ msg);
                Loger.recv(string.Format("接受：{0}", msg));

                //删除前n个字节
                m_recvBytes.RemoveRange(0, index + back.Length + 4);
            }
            else
                m_recvBytes.RemoveRange(0, index + 1); 
        }

        private void funCode5(List<byte> recvBytes, int index)
        {
            Back5 back = new Back5();
            if (m_recvBytes[index + back.Length + 3] == getCRC(recvBytes.ToArray(), index, back.Length + 3))
            {
                //判断不同的关键词，调用不同类的处理方法

                byte[] datas = new byte[back.Length];
                Array.Copy(m_recvBytes.ToArray(), index + 3, datas, 0, back.Length);

                //根据不同的功能码，获取不同的返回数据
                back.DataBytes = datas;

                string msg = string.Format("{0}: 体重系数:K={1},B={2}; 通道1：K={3}，B={4}; ; 通道2：K={5}，B={6}; 通道3：K={7}，B={8}; 通道4：K={9}，B={10}; 通道5：K={11}，B={12}; 通道6：K={13}，B={14};",
                                            DateTime.Now.ToString("HH:mm:ss.fff"), 
                                            back.WeightRatio.K, back.WeightRatio.B,
                                            back.Resistance1Ratio.K,back.Resistance1Ratio.B,
                                            back.Resistance2Ratio.K,back.Resistance2Ratio.B, 
                                            back.Resistance3Ratio.K,back.Resistance3Ratio.B, 
                                            back.Resistance4Ratio.K,back.Resistance4Ratio.B, 
                                            back.Resistance5Ratio.K,back.Resistance5Ratio.B, 
                                            back.Resistance6Ratio.K,back.Resistance6Ratio.B);
                logFile(msg);
                Loger.recv(string.Format("接受：{0}", msg));

                this.Invoke(new DelegatBack(back5Recieved), new object[] { back });
                m_recvBytes.RemoveRange(0, index + back.Length + 4);
            }
            else
                m_recvBytes.RemoveRange(0, index + 1);
        }
        private void back5Recieved(Back back)
        {
            //Back5 back5 = back as Back5;

            //txt5WeightK.Text = back5.WeightRatio.K.ToString();
            //txt5WeightB.Text = back5.WeightRatio.B.ToString();

            //txt5Resist1K.Text = back5.Resistance1Ratio.K.ToString();
            //txt5Resist1B.Text = back5.Resistance1Ratio.B.ToString();

            //txt5Resist2K.Text = back5.Resistance2Ratio.K.ToString();
            //txt5Resist2B.Text = back5.Resistance2Ratio.B.ToString();

            //txt5Resist3K.Text = back5.Resistance3Ratio.K.ToString();
            //txt5Resist3B.Text = back5.Resistance3Ratio.B.ToString();

            //txt5Resist4K.Text = back5.Resistance4Ratio.K.ToString();
            //txt5Resist4B.Text = back5.Resistance4Ratio.B.ToString();

            //txt5Resist5K.Text = back5.Resistance5Ratio.K.ToString();
            //txt5Resist5B.Text = back5.Resistance5Ratio.B.ToString();

            //txt5Resist6K.Text = back5.Resistance6Ratio.K.ToString();
            //txt5Resist6B.Text = back5.Resistance6Ratio.B.ToString();

        }

        private void funCode8(List<byte> recvBytes, int index)
        {
            Back8_2 back = new Back8_2();
            if (m_recvBytes[index + back.Length + 3] == getCRC(recvBytes.ToArray(), index, back.Length + 3))
            {
                byte[] datas = new byte[back.Length];
                Array.Copy(m_recvBytes.ToArray(), index + 3, datas, 0, back.Length);

                //根据不同的功能码，获取不同的返回数据
                back.DataBytes = datas;

                string msg = string.Format("{0}: 频段号,{1},体重,{2},阻抗1,{3},阻抗2,{4}",
                                             DateTime.Now.ToString("HH:mm:ss.fff"), back.频段号, back.体重, back.阻抗1, back.阻抗2);
                logFile(msg);
                Loger.recv(string.Format("接受：{0}", msg));

                this.Invoke(new DelegatBack(back8Recieved), new object[] { back });
                //删除前n个字节
                m_recvBytes.RemoveRange(0, index + back.Length + 4);
            }
            else
                //删除前n个字节
                m_recvBytes.RemoveRange(0, index + 1);
        }
        private void funCode8_3(List<byte> recvBytes, int index)
        {
            Back8_3 back = new Back8_3();
            if (true)
            {
                byte[] datas = new byte[back.Length];
                Array.Copy(m_recvBytes.ToArray(), index + 3, datas, 0, back.Length);

                //根据不同的功能码，获取不同的返回数据
                back.DataBytes = datas;

                //if (back.通道类型 == 2)
                {
                    string msg = string.Format("{0}: 频段号,{1},体重,{2},阻抗1,{3},阻抗2,{4},阻抗3,{5},阻抗4,{6},阻抗5,{7}",
                                                 DateTime.Now.ToString("HH:mm:ss.fff"), 
                                                 back.频段号, back.体重, back.阻抗1, back.阻抗2,
                                                 back.阻抗3, back.阻抗4, back.阻抗5);

                    logFile(msg);
                    Loger.recv(string.Format("接受：{0}", msg));

                    this.Invoke(new DelegatBack(back8Recieved), new object[] { back });
                }
                //删除前n个字节
                m_recvBytes.RemoveRange(0, index + back.Length + 3);
            }
            else
                //删除前n个字节
                m_recvBytes.RemoveRange(0, index + 1);
        }
        //

        private void funCodeA(List<byte> recvBytes, int index)
        {
            BackA back = new BackA();
            if (m_recvBytes[index + back.Length + 3] == getCRC(recvBytes.ToArray(), index, back.Length + 3))
           {
                byte[] datas = new byte[back.Length];
                Array.Copy(m_recvBytes.ToArray(), index + 3, datas, 0, back.Length);

                //根据不同的功能码，获取不同的返回数据
                back.DataBytes = datas;

                try 
	            {	        
		                string msg = string.Format("{0},体重K,{1},体重B,{2},5khz1K,{3},5khz1B,{4},5khz2K,{5},5khz2B,{6},250khz1K,{7},250khz1B,{8},250khz2K,{9},250khz2B,{10}",
                                                                                    DateTime.Now.ToString("HH:mm:ss.fff"),
                                                                                    back.体重K, back.体重B, 
                                                                                    back.阻抗5K1, back.阻抗5B1,
                                                                                    back.阻抗5K2, back.阻抗5B2,
                                                                                    back.阻抗250K1, back.阻抗250B1,
                                                                                    back.阻抗250K2,back.阻抗250B2
                                                                                    );
                    logFile(msg);
                    Loger.recv(string.Format("接受：{0}", msg));

                    this.Invoke(new DelegatBack(backARecieved), new object[] { back });
	            }
	            catch
	            {
	            }

                //删除前n个字节
                m_recvBytes.RemoveRange(0, index + back.Length + 4);
           }
            else
                //删除前n个字节
                m_recvBytes.RemoveRange(0, index + 1);
        }
        private void backARecieved(Back back)
        {
            //BackA backa = back as BackA;

            //txt体重k.Text = backa.体重K.ToString();
            //txt体重b.Text = backa.体重B.ToString();

            //txt5k1.Text = backa.阻抗5K1.ToString();
            //txt5b1.Text = backa.阻抗5B1.ToString();

            ////txt5b2
            //txt5k2.Text = backa.阻抗5K2.ToString();
            //txt5b2.Text = backa.阻抗5B2.ToString();

            //txt250k1.Text = backa.阻抗250K1.ToString();
            //txt250b1.Text = backa.阻抗250B1.ToString();

            //txt250k2.Text = backa.阻抗250K2.ToString();
            //txt250b2.Text = backa.阻抗250B2.ToString();

        }
        private void back8Recieved(Back back)
        {
            Back8_3 back8 = back as Back8_3;

            //txt8BaudId.Text = back8.频段号.ToString();
            //txt8Weight.Text = back8.体重.ToString();

            //txtWeight.Text = back8.体重.ToString();

            //if (back8.频段号 == 7)
            //{
            //    txt250ra.Text = back8.阻抗1.ToString();
            //    txt250la.Text = back8.阻抗2.ToString();
            //    txt250tr.Text = back8.阻抗3.ToString();
            //    txt250rl.Text = back8.阻抗4.ToString();
            //    txt250ll.Text = back8.阻抗5.ToString();

            //    caclPBF();
            //}
            //else if(back8.频段号 == 4)
            //{
            //    txt50ra.Text = back8.阻抗1.ToString();
            //    txt50la.Text = back8.阻抗2.ToString();
            //    txt50tr.Text = back8.阻抗3.ToString();
            //    txt50rl.Text = back8.阻抗4.ToString();
            //    txt50ll.Text = back8.阻抗5.ToString();
            //}

        }

        private void logFile(string msg)
        {
            sw.WriteLine(msg);
            sw.Flush();
        }
        private void backRecieved(Back back)
        {
            Back4 back4 = back as Back4;
            //txtBaud.Text = back4.频段号.ToString();
            //txtWeight.Text = back4.体重.ToString();
            //txtZukang1.Text = back4.阻抗1.ToString();
            //txtZukang2.Text = back4.阻抗2.ToString();
            //txtZukang3.Text = back4.阻抗3.ToString();
            //txtZukang4.Text = back4.阻抗4.ToString();
            //txtZukang5.Text = back4.阻抗5.ToString();
            //txtZukang6.Text = back4.阻抗6.ToString();
        }
        //遍历list
        private int findIndex(List<byte> dataList, byte value, int start)
        {
            int index = start;
            foreach (byte temp in dataList)
            {
                ++index;
                if (temp == value)
                {
                    break;
                }
                
            }
            return index;
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            try
            {
                textValue2Ini("LA", txtLAIm);
                textValue2Ini("RA", txtRAIm);
                textValue2Ini("TT", txtTTIm);
                textValue2Ini("LL", txtLLIm);
                textValue2Ini("RL", txtRLIm);

                textValue2Ini("LAP", txtLAP);
                textValue2Ini("RAP", txtRAP);
                textValue2Ini("TTP", txtTTP);
                textValue2Ini("LLP", txtLLP);
                textValue2Ini("RLP", txtRLP);

                cbxValue2Ini("com", cbxCom);
                cbxValue2Ini("baud", cboxBaud);

                sw.Close();
                fs.Close();
            }
            catch
            { }
        }


        private void cbxValue2Ini(string key, ComboBox cbx)
        {
            try
            {
                INI.WriteIni(key, cbx.Text);
            }
            catch (Exception ex)
            {

                Loger.send(ex.ToString());
            }

        }        

        private void textValue2Ini(string key, TextBox txt)
        {
            try
            {
                INI.WriteIni(key, txt.Text);
            }
            catch (Exception ex)
            {

                Loger.send(ex.ToString());
            }
 
        }
        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                m_serialPort.Close();
                m_serialPort = null;
                Loger.send("串口成功关闭");
            }
            catch 
            {

            }

        }

        private void btn02_Click(object sender, EventArgs e)
        {
            //try
            //{
            //    byte[] srcUserId = System.Text.Encoding.Default.GetBytes(txtUserID.Text);

            //    byte[] sendData = new byte[16];
            //    sendData[0] = 0x64;
            //    sendData[1] = 0x95;
            //    sendData[2] = 12;
            //    sendData[3] = 0x2;
            //    if (srcUserId.Length <= 8 && srcUserId.Length > 0)
            //        Array.Copy(srcUserId, 0, sendData, 4, srcUserId.Length);
            //    else if (srcUserId.Length > 8)
            //        Array.Copy(srcUserId, 0, sendData, 4, 8);
            //    else
            //        Loger.send("用户编号为空！");

            //    sendData[12] = Convert.ToByte(txtHeight.Text);
            //    sendData[13] = m_gender;
            //    sendData[14] = Convert.ToByte(txtAge.Text);
            //    sendData[15] = getXor(sendData, 0, sendData.Length - 1);

            //    SendCommand(sendData);

            //    Loger.send(logBytes(sendData));
            //}
            //catch (Exception ex)
            //{
            //    Loger.send(ex.Message);
            //}
        }

        private void cbxGender_SelectedValueChanged(object sender, EventArgs e)
        {
            //if (cbxGender.Text == "男")
            //{
            //    m_gender = 0;
            //}
            //else
            //    m_gender = 1;
        }

        private string logBytes(byte[] bytes)
        {
            string logMsg = string.Empty;
            foreach (byte b in bytes)
            {
                logMsg += b.ToString("x2");
                logMsg += "  ";
            }
            return logMsg;
        }

        private void btn05_Click(object sender, EventArgs e)
        {
            try
            {
                if (m_serialPort != null && m_serialPort.IsOpen)
                {
                    byte[] sendData = new byte[5];
                    sendData[0] = 0x64;
                    sendData[1] = 0x95;
                    sendData[2] = 1;
                    sendData[3] = 0x05;
                    sendData[4] = getCRC(sendData, 0, sendData.Length - 1);

                    string msg = string.Empty;
                    foreach (byte temp in sendData)
                    {
                        msg += temp.ToString("x2");
                        msg += "  ";
                    }
                    Loger.send(string.Format("发送：{0}", msg));

                    int n = SendCommand(sendData);
                }
                else
                    Loger.send("串口未打开");
            }
            catch (Exception ex)
            {
                Loger.send("打开串口失败！");
                Loger.send(ex.Message);
            }
        }

        private void btn08_Click(object sender, EventArgs e)
        {
            try
            {
                if (m_serialPort != null && m_serialPort.IsOpen)
                {
                    byte[] sendData = new byte[6];
                    sendData[0] = 0x64;
                    sendData[1] = 0x95;
                    sendData[2] = 2;
                    sendData[3] = 0x08;
                    sendData[4] = m_baud08;
                    sendData[5] = getCRC(sendData, 0, sendData.Length - 1);

                    string msg = string.Empty;
                    foreach (byte temp in sendData)
                    {
                        msg += temp.ToString("x2");
                        msg += "  ";
                    }
                    Loger.send(string.Format("发送：{0}", msg));

                    int n = SendCommand(sendData);
                }
                else
                    Loger.send("串口未打开");
            }
            catch (Exception ex)
            {
                Loger.send("打开串口失败！");
                Loger.send(ex.Message);
            }
        }

        private void cbxBaud08_SelectedValueChanged(object sender, EventArgs e)
        {
            //switch (cbxBaud08.Text.Trim())
            //{
            //    case "1khz":
            //        m_baud08 = 0;
            //        break;
            //    case "5khz":
            //        m_baud08 = 1;
            //        break;
            //    case "10khz":
            //        m_baud08 = 2;
            //        break;
            //    case "20khz":
            //        m_baud08 = 3;
            //        break;
            //    case "50khz":
            //        m_baud08 = 4;
            //        break;
            //    case "100khz":
            //        m_baud08 = 5;
            //        break;
            //    case "200khz":
            //        m_baud08 = 6;
            //        break;
            //    case "250khz":
            //        m_baud08 = 7;
            //        break;
            //    case "500khz":
            //        m_baud08 = 8;
            //        break;
            //    case "1000khz":
            //        m_baud08 = 9;
            //        break;
            //}
        }

        private void cbxBaud_SelectedValueChanged_1(object sender, EventArgs e)
        {
            //switch (cbxBaud.Text.Trim())
            //{
            //    case "1khz":
            //        m_baud = 0;
            //        break;
            //    case "5khz":
            //        m_baud = 1;
            //        break;
            //    case "10khz":
            //        m_baud = 2;
            //        break;
            //    case "20khz":
            //        m_baud = 3;
            //        break;
            //    case "50khz":
            //        m_baud = 4;
            //        break;
            //    case "100khz":
            //        m_baud = 5;
            //        break;
            //    case "200khz":
            //        m_baud = 6;
            //        break;
            //    case "250khz":
            //        m_baud = 7;
            //        break;
            //    case "500khz":
            //        m_baud = 8;
            //        break;
            //    case "1000khz":
            //        m_baud = 9;
            //        break;
            //}
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //try
            //{
            //    byte tongdao = 0;
            //    #region tongdao
            //    switch (cbxTongdao.Text)
            //    {
            //        case "体重":
            //            tongdao = 0;
            //            break;
            //        case "通道1":
            //            tongdao = 1;
            //            break;
            //        case "通道2":
            //            tongdao = 2;
            //            break;
            //        case "通道3":
            //            tongdao = 3;
            //            break;
            //        case "通道4":
            //            tongdao = 4;
            //            break;
            //        case "通道5":
            //            tongdao = 5;
            //            break;
            //        case "通道6":
            //            tongdao = 6;
            //            break;
            //        default:
            //            break;
            //    }
            //    #endregion

            //    byte[] sendData = new byte[14];
            //    sendData[0] = 0x64;
            //    sendData[1] = 0x95;
            //    sendData[2] = 10;
            //    sendData[3] = 0x06;
            //    sendData[4] = tongdao;

            //    float k = Convert.ToSingle(txtK.Text);
            //    byte[] kbytes = BitConverter.GetBytes(k);
            //    Array.Copy(kbytes, 0, sendData, 5, 4);

            //    float b = Convert.ToSingle(txtB.Text);
            //    byte[] bbytes = BitConverter.GetBytes(b);
            //    Array.Copy(bbytes, 0, sendData, 9, 4);

            //    sendData[13] = getXor(sendData, 0, sendData.Length - 1);

            //    string msg = string.Empty;
            //    foreach (byte temp in sendData)
            //    {
            //        msg += temp.ToString("x2");
            //        msg += "  ";
            //    }
            //    Loger.send(string.Format("发送：{0}", msg));

            //    if (m_serialPort != null && m_serialPort.IsOpen)
            //        SendCommand(sendData);
            //    else Loger.send("串口未打开");
            //}
            //catch (Exception ex)
            //{
            //    Loger.send(ex.Message);
            //}
        }

        private void cbxWork_SelectedValueChanged(object sender, EventArgs e)
        {
            //switch (cbxWork.Text)
            //{
            //    case "校准模式":
            //        m_isWorkStyle = false;
            //        break;
            //    case "正常模式":
            //        m_isWorkStyle = true;
            //        break;
            //    default:
            //        break;
            //}
        }

        private void btnChangeWork_Click(object sender, EventArgs e)
        {
            try
            {
                byte[] sendData = new byte[10];
                sendData[0] = 0x64;
                sendData[1] = 0x95;
                sendData[2] = 6;
                sendData[3] = 0x01;
                sendData[4] =(byte)(m_isWorkStyle==true?1:0);
                sendData[5] = 0x45;
                sendData[6] = 0x85;
                sendData[7] = 0x72;
                sendData[8] = 0x01;
          
                sendData[9] = getCRC(sendData, 0, sendData.Length - 1);

                string msg = string.Empty;
                foreach (byte temp in sendData)
                {
                    msg += temp.ToString("x2");
                    msg += "  ";
                }
                Loger.send(string.Format("发送：{0}", msg));

                if (m_serialPort != null && m_serialPort.IsOpen)
                    SendCommand(sendData);
                else Loger.send("串口未打开");
            }
            catch (Exception ex)
            {
                Loger.send(ex.Message);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                byte[] sendData = new byte[5];
                sendData[0] = 0x64;
                sendData[1] = 0x95;
                sendData[2] = 1;
                sendData[3] = 0x09;
                sendData[4] = getCRC(sendData, 0, sendData.Length - 1);

                string msg = string.Empty;
                foreach (byte temp in sendData)
                {
                    msg += temp.ToString("x2");
                    msg += "  ";
                }
                Loger.send(string.Format("发送：{0}", msg));

                if (m_serialPort != null && m_serialPort.IsOpen)
                    SendCommand(sendData);
                else Loger.send("串口未打开");
            }
            catch (Exception ex)
            {
                Loger.send(ex.Message);
            }
        }

        private void btn07_Click(object sender, EventArgs e)
        {
            try
            {
                byte[] sendData = new byte[5];
                sendData[0] = 0x64;
                sendData[1] = 0x95;
                sendData[2] = 1;
                sendData[3] = 0x07;
                sendData[4] = getCRC(sendData, 0, sendData.Length - 1);

                string msg = string.Empty;
                foreach (byte temp in sendData)
                {
                    msg += temp.ToString("x2");
                    msg += "  ";
                }
                Loger.send(string.Format("发送：{0}", msg));

                if (m_serialPort != null && m_serialPort.IsOpen)
                    SendCommand(sendData);
                else Loger.send("串口未打开");
            }
            catch (Exception ex)
            {
                Loger.send(ex.Message);
            }
        }

        private void funCode7(List<byte> recvBytes, int index)
        {
            Back7 back = new Back7();
            if (m_recvBytes[index + back.Length + 3] == getCRC(recvBytes.ToArray(), index, back.Length + 3))
            {
                //判断不同的关键词，调用不同类的处理方法

                byte[] datas = new byte[back.Length];
                Array.Copy(m_recvBytes.ToArray(), index + 3, datas, 0, back.Length);

                //根据不同的功能码，获取不同的返回数据
                back.DataBytes = datas;

                string msg = string.Format("{0}: 数据包号,{1},用户编号,{2},体重值,{3},脂肪重,{4},脂肪率,{5},总体水,{6},细胞外液,{7},肌肉率,{8},蛋白质,{9},矿物质,{10},基础代谢率,{11}",
                                             DateTime.Now.ToString("HH:mm:ss.fff"), back.数据包号, back.用户编号, back.体重值, back.脂肪重, back.脂肪率, back.总体水, back.细胞外液, back.肌肉率,
                                                                                                                 back.蛋白质,back.矿物质,back.基础代谢率);
                logFile(msg);
                Loger.recv(string.Format("接受：{0}", msg));

                this.Invoke(new DelegatBack(back7Recieved), new object[] { back });
                //删除前n个字节
                m_recvBytes.RemoveRange(0, index + back.Length + 4);
            }
            else
                //删除前n个字节
                m_recvBytes.RemoveRange(0, index + 1);
        }

        private void back7Recieved(Back back)
        {
            Back7 back7 = back as Back7;

            //D1：数据包号：1字节，0~255循环递增，以示数据的连续性。
            //txt数据包号.Text=back7.数据包号.ToString();
            ////D2：用户编号：8字节字符串
            //txt用户编号.Text = back7.用户编号.ToString();
            ////D3：体重值：4字节，float型，单位为kg
            //txt体重值.Text = back7.体重值.ToString();
            ////D4：脂肪重：4字节，float型，单位为kg
            //txt脂肪重.Text = back7.脂肪重.ToString();
            ////D5：脂肪率：4字节，float型，单位为%（如56.3代表56.3%）
            //txt脂肪率.Text = back7.脂肪率.ToString();
            ////D6：总体水：4字节，float型，单位为kg
            //txt总体水.Text = back7.总体水.ToString();
            ////D7：细胞外液：4字节，float型，单位为kg
            //txt细胞外液.Text = back7.细胞外液.ToString();
            ////D8：肌肉率：4字节，float型，单位为%（如56.3代表56.3%）
            //txt肌肉率.Text = back7.肌肉率.ToString();
            ////D9：蛋白质：4字节，float型，单位为kg
            //txt蛋白质.Text = back7.蛋白质.ToString();
            ////D10：矿物质：4字节，float型，单位为kg
            //txt矿物质.Text = back7.矿物质.ToString();
            ////D11：基础代谢率：4字节，float型，单位为kcal/d  
            //txt基础代谢率.Text = back7.基础代谢率.ToString();

        }

        private void btnUserDefine_Click(object sender, EventArgs e)
        {
            //try
            //{
            //    string inputMsg = txtInput.Text;
            //    int dataLength=m_bytes.Length;

            //    byte[] sendData = new byte[4+dataLength+2];
            //    sendData[0] = 0x64;
            //    sendData[1] = 0x95;
            //    sendData[2] = (byte)(dataLength+2);
            //    sendData[3] =0x0c;
            //    if (txtDataType.Text.Trim().Length > 0)
            //    {
            //        sendData[4] = Convert.ToByte(txtDataType.Text.Trim(), 16);
            //    }
            //    else
            //    {
            //        Loger.send("数据类型未选择");
            //        return;
            //    }
            //    Array.Copy(m_bytes,0, sendData, 5, m_bytes.Length);
                
            //    sendData[sendData.Length-1] = getXor(sendData, 0, sendData.Length - 1);

            //    string msg = string.Empty;
            //    foreach (byte temp in sendData)
            //    {
            //        msg += temp.ToString("x2");
            //        msg += "  ";
            //    }
            //    Loger.send(string.Format("发送：{0}", msg));
            //    logFile(string.Format(DateTime.Now.ToString("HH:mm:ss.fff")+" 发送：{0}", msg));

            //    if (m_serialPort != null && m_serialPort.IsOpen)
            //        SendCommand(sendData);
            //    else Loger.send("串口未打开");
            //}
            //catch (Exception ex)
            //{
            //    Loger.send(ex.Message);
            //}
        }
        //处理字节
        private byte[] dealBytes()
        {
            List<byte> bList = new List<byte>(); 
            //string[] inputs = txtInput.Text.Split(new char[]{',', ' '});

            //foreach (string item in inputs)
            //{
            //        Match m = Regex.Match(item, "[\\da-fA_F]{2}");
            //        if (m.Length > 0)
            //        {
            //            try
            //            {
            //                byte temp = (byte)Convert.ToInt32(m.Value, 16);
            //                bList.Add(temp);
            //            }
            //            catch
            //            {
            //                Loger.send("字节有错误");
            //                break;
            //            }
            //        }
            //}  
            return bList.ToArray();
        }
        //处理float
        private byte[] dealFloat()
        {
            List<byte> bList = new List<byte>();
            //string[] inputs = txtInput.Text.Split(new char[] { ',', ' ' });

            //foreach (string item in inputs)
            //{
            //    Match m = Regex.Match(item, "[-\\d.]+");
            //    if (m.Length > 0)
            //    {
            //        try
            //        {
            //            float temp = (float)Convert.ToSingle(m.Value);
            //            byte[] tempbytes=BitConverter.GetBytes(temp);
            //            foreach (byte itemb in tempbytes)
            //            {
            //                bList.Add(itemb);
            //            }
            //        }
            //        catch
            //        {
            //            Loger.send("输入内容有错误");
            //            break;
            //        }
            //    }

            //}
            return bList.ToArray();
        }
        //处理ushort
        private byte[] dealUshort()
        {
            List<byte> bList = new List<byte>();
            //string[] inputs = txtInput.Text.Split(new char[] { ',', ' ' });

            //foreach (string item in inputs)
            //{
            //    Match m = Regex.Match(item, "[\\d]+");
            //    if (m.Length > 0)
            //    {
            //        try
            //        {
            //            UInt16 temp = (UInt16)Convert.ToUInt16(m.Value);
            //            byte[] tempbytes = BitConverter.GetBytes(temp);
            //            foreach (byte itemb in tempbytes)
            //            {
            //                bList.Add(itemb);
            //            }
            //        }
            //        catch
            //        {
            //            Loger.send("字节有错误");
            //            break;
            //        }
            //    }

            //}
            return bList.ToArray();
        }

        //处理short
        private byte[] dealShort()
        {
            List<byte> bList = new List<byte>();
            //string[] inputs = txtInput.Text.Split(new char[] { ',', ' ' });

            //foreach (string item in inputs)
            //{
            //    Match m = Regex.Match(item, "[\\d]+");
            //    if (m.Length > 0)
            //    {
            //        try
            //        {
            //            Int16 temp = (Int16)Convert.ToUInt16(m.Value);
            //            byte[] tempbytes = BitConverter.GetBytes(temp);
            //            foreach (byte itemb in tempbytes)
            //            {
            //                bList.Add(itemb);
            //            }
            //        }
            //        catch
            //        {
            //            Loger.send("字节有错误");
            //            break;
            //        }
            //    }
            //}
            return bList.ToArray();
        }
        private void txtInput_TextChanged(object sender, EventArgs e)
        {
            inputTextChanged();
        }

        private void inputTextChanged()
        {
            //string inputMsg = txtInput.Text;

            //switch (txtDataType.Text)
            //{
            //    case "03":
            //        //字节
            //        m_bytes = dealBytes();
            //        break;
            //    case "01":
            //        //ushort
            //        m_bytes = dealUshort();
            //        break;
            //    case "02":
            //        //short
            //        m_bytes = dealShort();
            //        break;
            //    case "00":
            //        //float
            //        m_bytes = dealFloat();
            //        break;
            //}

            //if (m_bytes != null && m_bytes.Length > 0)
            //    txtLength.Text = (m_bytes.Length + 2).ToString("x2");//mc.Count.ToString("x2");
        }
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            cbxTypeChanged();

            inputTextChanged();
        }
        private void cbxTypeChanged()
        {
            //switch (cbxDataType.Text)
            //{
            //    case "byte":
            //        txtDataType.Text = "03";
            //        break;

            //    case "ushort":
            //        txtDataType.Text = "01";
            //        break;

            //    case "short":
            //        txtDataType.Text = "02";
            //        break;

            //    case "float":
            //        txtDataType.Text = "00";
            //        break;
            //}
        }

        private void groupBox11_Enter(object sender, EventArgs e)
        {

        }

        private void btnCaclPBF_Click(object sender, EventArgs e)
        {
            caclPBF();
        }
        private void caclPBF()
        {
            //PBF	体脂肪百分比
            //float PBF = 0;
            //try
            //{
            //    if (txtHeight.Text.Trim().Length <= 0)
            //    {
            //        MessageBox.Show("身高数据不能为空");
            //        return;
            //    }
            //    if (txtAge.Text.Trim().Length <= 0)
            //    {
            //        MessageBox.Show("年龄数据不能为空");
            //        return;
            //    }
            //    if (txtWeight.Text.Trim().Length <= 0)
            //    {
            //        MessageBox.Show("体重数据不能为空");
            //        return;
            //    }

            //    int height = Int32.Parse(txtHeight.Text);
            //    int age = Int32.Parse(txtAge.Text);

            //    float ra50 = float.Parse(txt50ra.Text);
            //    float la50 = float.Parse(txt50la.Text);
            //    float tr50 = float.Parse(txt50tr.Text);
            //    float rl50 = float.Parse(txt50rl.Text);
            //    float ll50 = float.Parse(txt50ll.Text);

            //    float ra250 = float.Parse(txt250ra.Text);
            //    float la250 = float.Parse(txt250la.Text);
            //    float tr250 = float.Parse(txt250tr.Text);
            //    float rl250 = float.Parse(txt250rl.Text);
            //    float ll250 = float.Parse(txt250ll.Text);

            //    float weight = float.Parse(txtWeight.Text);

            //    BIA bia = new BIA();
            //    bia.set_impedance(height,weight,1,age,1,ra50,la50,tr50,rl50,ll50,ra250,la250,tr250,rl250,ll250);
            //    PBF= (float)bia.getPBF();

            //    txtPBF.Text = PBF.ToString();
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.ToString());
            //}
            //try
            //{
            //    //发送数据
            //    byte[] sendData = new byte[8];
            //    sendData[0] = 0x64;
            //    sendData[1] = 0x95;
            //    sendData[2] = 5;
            //    sendData[3] = 0x17;

            //    byte[] PBFbyte = new byte[4];
            //    PBFbyte = BitConverter.GetBytes(PBF);
            //    Array.Copy(PBFbyte, 0, sendData, 4, 4);
            //    sendData[7] = getXor(sendData, 0, sendData.Length - 1);

            //    SendCommand(sendData);

            //    Loger.send(logBytes(sendData));
            //}
            //catch (Exception ex)
            //{
            //    Loger.send(ex.Message);
            //}
        }
        private void btnClearData_Click(object sender, EventArgs e)
        {
            //richboxImpedance.Clear();

            //txt250ra.Text = string.Empty;
            //txt250la.Text = string.Empty;
            //txt250tr.Text = string.Empty;
            //txt250rl.Text = string.Empty;
            //txt250ll.Text = string.Empty;

            //txt50ra.Text = string.Empty;
            //txt50la.Text = string.Empty;
            //txt50tr.Text = string.Empty;
            //txt50rl.Text = string.Empty;
            //txt50ll.Text = string.Empty;

            //txtPBF.Text = string.Empty;
            //txtWeight.Text = string.Empty;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            richBoxSend.Clear();
        }

        private void btnWeight_Click(object sender, EventArgs e)
        {
                              //02  53  0e  0a  4c  43  4c  1b  2b  03 
            btnWeight.TextColor = Color.Red;
            textWeight.Text = "";
            send(new byte[] { 0x02, 0x53, 0x0e, 0x0a, 0x4c, 0x43, 0x4c, 0x1b, 0x2b, 0x03 });
        }

        private void btn_Click(object sender, EventArgs e)
        {
            richboxReceive.Clear();
        }

        private void send(byte[] datas)
        {
            try
            {
                if (datas == null || datas.Length <= 0)
                {
                    Loger.send("发送数据为空");
                    return;
                }

                Loger.send(logBytes(datas));

                // Encrypt messag //

                //byte key = (byte)(new Random().Next(128, 255));
                //datas[1] = key;
                //codec(datas);
                //datas[datas.Length - 2] = getCRC(datas, 1, datas.Length - 3);

                SendCommand(datas);

            }
            catch (Exception ex)
            {
                Loger.send(ex.Message);
            }         
        }

        private void btnWeight100Clear_Click(object sender, EventArgs e)
        {
            btnCalibration100.TextColor = Color.Red;
            m_calibariton_type = 100;
            send(new byte[] { 0x2 , 0x53 , 0xe , 0xa , 0x4c , 0x43 , 0x43 , 0x1b , 0x22 , 0x3 });
            Thread.Sleep(5000);
            send(new byte[] { 0x2 , 0x53 , 0xe , 0xa , 0x4c , 0x43 , 0x34 , 0x1b , 0x13 , 0x3 });
            Thread.Sleep(5000);
            send(new byte[] { 0x2 , 0x53 , 0xe , 0xa , 0x4c , 0x43 , 0x58 , 0x1b , 0x37 , 0x3 });
        }

        private void btnWeight75Clear_Click(object sender, EventArgs e)
        {
            btnCalibration75.TextColor = Color.Red;
            m_calibariton_type = 75;
            send(new byte[] { 0x2 , 0x53 , 0xe , 0xa , 0x4c , 0x43 , 0x43 , 0x1b , 0x22 , 0x3  });

            Thread.Sleep(5000);
            send(new byte[] { 0x2 , 0x53 , 0xe , 0xa , 0x4c , 0x43 , 0x33 , 0x1b , 0x12 , 0x3 });
            Thread.Sleep(5000);
            send(new byte[] { 0x2 , 0x53 , 0xe , 0xa , 0x4c , 0x43 , 0x58 , 0x1b , 0x37 , 0x3 });
        }

        private void btnWeight50Clear_Click(object sender, EventArgs e)
        { 
            btnCalibration50.TextColor = Color.Red;
            m_calibariton_type = 50;
            send(new byte[] { 0x2 , 0x53 , 0xe , 0xa , 0x4c , 0x43 , 0x43 , 0x1b , 0x22 , 0x3 });
            Thread.Sleep(5000);
            send(new byte[] { 0x2 , 0x53 , 0xe , 0xa , 0x4c , 0x43 , 0x32 , 0x1b , 0x11 , 0x3 });
            Thread.Sleep(5000);
            send(new byte[] { 0x2 , 0x53 , 0xe , 0xa , 0x4c , 0x43 , 0x58 , 0x1b , 0x37 , 0x3 });
        }

        private void btnWeight25Clear_Click(object sender, EventArgs e)
        { 
            btnCalibration25.TextColor = Color.Red;
            m_calibariton_type = 25;
            send(new byte[] { 0x2 , 0x53 , 0xe , 0xa , 0x4c , 0x43 , 0x43 , 0x1b , 0x22 , 0x3 });

            Thread.Sleep(5000);
            send(new byte[] { 0x2 , 0x53 , 0xe , 0xa , 0x4c , 0x43 , 0x31 , 0x1b , 0x10 , 0x3 });
            Thread.Sleep(5000);
            send(new byte[] { 0x2 , 0x53 , 0xe , 0xa , 0x4c , 0x43 , 0x58 , 0x1b , 0x37 , 0x3 });
        }

        private void btnWeight0Clear_Click(object sender, EventArgs e)
        {
            btnCalibration0.TextColor = Color.Red;
            m_calibariton_type = 0;
            send(new byte[] { 0x2 , 0x53 , 0xe , 0xa , 0x4c,  0x43 , 0x43 , 0x1b , 0x22 , 0x3  });

            Thread.Sleep(5000);
            send(new byte[] { 0x2 , 0x53 , 0xe , 0xa , 0x4c , 0x43 , 0x30 , 0x1b , 0xf , 0x3 });
            Thread.Sleep(5000);
            send(new byte[] { 0x2 , 0x53 , 0xe , 0xa , 0x4c , 0x43 , 0x58 , 0x1b , 0x37 , 0x3 });
        }

      //  private void btnCalibration_Click(object sender, EventArgs e)
       // {
       //     btnCalibration0.TextColor = Color.Red;
       //     m_calibariton_type = 10;
       //     int sleepTime = 5000;

       //     send(new byte[] { 0x02, 0x53, 0x0e, 0x0a, 0x4c, 0x43, 0x43, 0x1b, 0x22, 0x03 });
       //     Thread.Sleep(sleepTime);
       //     send(new byte[] { 0x02, 0x53, 0x0e, 0x0a, 0x4c, 0x43, 0x31, 0x1b, 0x10, 0x03 });
      //      Thread.Sleep(sleepTime);
       //     send(new byte[] { 0x02, 0x53, 0x0e, 0x0a, 0x4c, 0x43, 0x53, 0x1b, 0x32, 0x03 });
       //     Thread.Sleep(sleepTime);
        //    send(new byte[] { 0x02, 0x53, 0x0e, 0x0a, 0x4c, 0x43, 0x58, 0x1b, 0x37, 0x03 });
       // }

        private void btnCDR0_Click(object sender, EventArgs e)
        {
            m_counter = 1;
            btnCDR0.TextColor = Color.Red;
            send(new byte[] { 0x02, 0x53, 0x0f, 0x0a, 0x43, 0x44, 0x52, 0x30, 0x1b, 0x1a, 0x03 });           
        }

        private void btnCDR1_Click(object sender, EventArgs e)
        {
            m_counter = 1;
            btnCDR1.TextColor = Color.Red;
            send(new byte[] { 0x02, 0x53, 0x0f, 0x0a, 0x43, 0x44, 0x52, 0x31, 0x1b, 0x1b, 0x03 });
        }

        private void btnCDR2_Click(object sender, EventArgs e)
        {
            m_counter = 1;
            btnCDR2.TextColor = Color.Red;
            send(new byte[] { 0x02, 0x53, 0x0f, 0x0a, 0x43, 0x44, 0x52, 0x32, 0x1b, 0x1c, 0x03 });
        }

        private void btnM1_Click(object sender, EventArgs e)
        {
            m_counter = 1;
            btnM1.TextColor = Color.Red;
            m_ReadStep = 1;
            // send M1
            send(new byte[] { 0x02, 0x53, 0x13, 0x0a, 0x4d, 0x31, 0x30, 0x31, 0x31, 0x31, 0x31, 0x30, 0x1b, 0x37, 0x03 });
        }

        private void btnReadImpedence_Click(object sender, EventArgs e)
        {
            btnReadImpedence.TextColor = Color.Red;
            clearImpedance();

            m_ReadStep = 3;
            // send CC0s
            send(new byte[] { 0x02, 0x53, 0x13, 0x0a, 0x43, 0x43, 0x30, 0x31, 0x31, 0x31, 0x31, 0x30, 0x1b, 0x3f, 0x03 });
        }
        private void btnMeasurephase_Click(object sender, EventArgs e)
        {
            //btnMeasurephase.TextColor = Color.Red;
            //clearPhase();

            //m_ReadStep = 1;
            // send PHA
            send(new byte[] { 0x02, 0x53, 0x0d, 0x0a, 0x45, 0x33, 0x1b, 0x47, 0x03 });
        }
        private void radiBtnHex_CheckedChanged(object sender, EventArgs e)
        {
            m_isAscii = false;
        }

        private void radiobtnAscii_CheckedChanged(object sender, EventArgs e)
        {
            m_isAscii = true;
        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            byte a = (byte)getrandom(0xf1);
            byte b =(byte) (a^ 0xb0);
            byte c = getrandom(b);
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void groupPanel1_Click(object sender, EventArgs e)
        {

        }

        private void buttonItem1_Click(object sender, EventArgs e)
        {

        }

        private void label8_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click_1(object sender, EventArgs e)
        {

        }

        private void labelLA5_TextChanged(object sender, EventArgs e)
        {
            upLabelBackGround(sender, m_LAStand, m_LAP);
        }

        private void upLabelBackGround(object sender, float stand, float percent)
        {
            getValueFromUI();

            try
            {
                LabelX label = (LabelX)sender;

                if (label.Text.Trim().Length <= 0) return;

                float im = float.Parse(((LabelX)sender).Text);

            float temp = (im - stand) / stand;
            if (Math.Abs(temp) <= percent/100)
            {
                label.BackColor = Color.Gainsboro;

            }
            else
            {
                label.BackColor = Color.Red;
            }
            }
            catch (Exception ex)
            {
                Loger.send(ex.ToString());
            }
        }

        private void getValueFromUI()
        {
            m_LAStand = getValue(txtLAIm);
            m_RAStand = getValue(txtRAIm);
            m_TTStand = getValue(txtTTIm);
            m_LLStand = getValue(txtLLIm);
            m_RLStand = getValue(txtRLIm);

            m_LAP = getValue(txtLAP);
            m_RAP = getValue(txtRAP);
            m_TTP = getValue(txtTTP);
            m_LLP = getValue(txtLLP);
            m_RLP = getValue(txtRLP);
        }

        private float getValue(TextBox txt)
        {
            if (txt.Text.Trim().Length > 0)
            {
                try
                {
                    return  float.Parse(txt.Text);
                }
                catch (Exception ex)
                {
                    Loger.send(ex.ToString());
                }
            }
            return -1;
        }

        private void labelLA50_TextChanged(object sender, EventArgs e)
        {
            upLabelBackGround(sender, m_LAStand, m_LAP);

        }

        private void labelLA250_TextChanged(object sender, EventArgs e)
        {
            upLabelBackGround(sender, m_LAStand, m_LAP);
        }

        private void labelLA500_TextChanged(object sender, EventArgs e)
        {
            upLabelBackGround(sender, m_LAStand, m_LAP);
        }

        private void labelRA5_TextChanged(object sender, EventArgs e)
        {
            upLabelBackGround(sender, m_RAStand, m_RAP);
        }

        private void labelRA50_TextChanged(object sender, EventArgs e)
        {
            upLabelBackGround(sender, m_RAStand, m_RAP);
        }

        private void labelRA250_TextChanged(object sender, EventArgs e)
        {
            upLabelBackGround(sender, m_RAStand, m_RAP);
        }

        private void labelRA500_TextChanged(object sender, EventArgs e)
        {
            upLabelBackGround(sender, m_RAStand, m_RAP);
        }

        private void labelTT5_TextChanged(object sender, EventArgs e)
        {
            upLabelBackGround(sender, m_TTStand, m_TTP);
        }

        private void labelTT50_TextChanged(object sender, EventArgs e)
        {
            upLabelBackGround(sender, m_TTStand, m_TTP);
        }

        private void labelTT250_TextChanged(object sender, EventArgs e)
        {
            upLabelBackGround(sender, m_TTStand, m_TTP);
        }

        private void labelTT500_TextChanged(object sender, EventArgs e)
        {
            upLabelBackGround(sender, m_TTStand, m_TTP);
        }

        private void labelLL5_TextChanged(object sender, EventArgs e)
        {
            upLabelBackGround(sender, m_LLStand, m_LLP);
        }

        private void labelLL50_TextChanged(object sender, EventArgs e)
        {
            upLabelBackGround(sender, m_LLStand, m_LLP);
        }

        private void labelLL250_TextChanged(object sender, EventArgs e)
        {
            upLabelBackGround(sender, m_LLStand, m_LLP);
        }

        private void labelLL500_TextChanged(object sender, EventArgs e)
        {
            upLabelBackGround(sender, m_LLStand, m_LLP);
        }

        private void labelRL5_TextChanged(object sender, EventArgs e)
        {
            upLabelBackGround(sender, m_RLStand, m_RLP);
        }

        private void labelRL50_TextChanged(object sender, EventArgs e)
        {
            upLabelBackGround(sender, m_RLStand, m_RLP);
        }

        private void labelRL250_TextChanged(object sender, EventArgs e)
        {
            upLabelBackGround(sender, m_RLStand, m_RLP);
        }

        private void labelRL500_TextChanged(object sender, EventArgs e)
        {
            upLabelBackGround(sender, m_RLStand, m_RLP);
        }

        private void buttonX2_Click(object sender, EventArgs e)
        {
            send(new byte[] { 0x02, 0x53, 0x0e, 0x0a, 0x4c, 0x43, 0x43, 0x1b, 0x22, 0x03 });
        }

        private void clearImpedance()
        {
            labelLA5.Text = "";
            labelLA50.Text = "";
            labelLA250.Text = "";
            labelLA500.Text = "";
            labelLA1K.Text = "";
            labelLA1M.Text = "";

            labelRA5.Text = "";
            labelRA50.Text = "";
            labelRA250.Text = "";
            labelRA500.Text = "";
            labelRA1K.Text = "";
            labelRA1M.Text = "";

            labelTT5.Text = "";
            labelTT50.Text = "";
            labelTT250.Text = "";
            labelTT500.Text = "";
            labelTT1K.Text = "";
            labelTT1M.Text = "";

            labelLL5.Text = "";
            labelLL50.Text = "";
            labelLL250.Text = "";
            labelLL500.Text = "";
            labelLL1K.Text = "";
            labelLL1M.Text = "";

            labelRL5.Text = "";
            labelRL50.Text = "";
            labelRL250.Text = "";
            labelRL500.Text = "";
            labelRL1K.Text = "";
            labelRL1M.Text = "";

            changeBackGround(new LabelX[]{
            labelLA5,
            labelLA50,
            labelLA250,
            labelLA500,
            labelLA1K,
            labelLA1M,

            labelRA5,
            labelRA50,
            labelRA250,
            labelRA500,
            labelRA1K,
            labelRA1M,

            labelTT5,
            labelTT50,
            labelTT250,
            labelTT500,
            labelTT1K,
            labelTT1M,

            labelLL5,
            labelLL50,
            labelLL250,
            labelLL500,
            labelLL1K,
            labelLL1M,

            labelRL5,
            labelRL50,
            labelRL250,
            labelRL500,
            labelRL1K,
            labelRL1M

            });
        }
        private void btnClearImpedance_Click(object sender, EventArgs e)
        {
            clearImpedance();
        }

        private void changeBackGround(LabelX[] labels)
        {
            foreach (var item in labels)
            {
                item.BackColor = Color.Gainsboro;
            }
        }

        private void cboxBaud_SelectedIndexChanged(object sender, EventArgs e)
        {


        }

        private void button4_Click_2(object sender, EventArgs e)
        {
            send(new byte[] { 0x02, 0x53, 0x0e, 0x0a, 0x4c, 0x43, 0x43, 0x1b, 0x22, 0x03 });

        }

        private void button5_Click_1(object sender, EventArgs e)
        {


            send(new byte[] { 0x02, 0x53, 0x0e, 0x0a, 0x4c, 0x43, 0x30, 0x1b, 0x0F, 0x03 });

        }

        private void btnNocalphase_Click(object sender, EventArgs e)
        {

            send(new byte[] { 0x02, 0x53, 0x0d, 0x0a, 0x49, 0x57, 0x1b, 0x2f, 0x03 });
        }

        private void button6_Click(object sender, EventArgs e)
        {
    
            send(new byte[] { 0x02, 0x53, 0x0e, 0x0a, 0x4c, 0x43, 0x58, 0x1b, 0x37, 0x03 });
        }

        private void buttonX1_Click(object sender, EventArgs e)
        {
            ButtonX[] buttons = new ButtonX[] { btnCalibration0, btnCalibration0, btnCDR2, btnCDR1, btnCDR0, btnWeight, btnM1, btnReadImpedence };

            foreach (var item in buttons)
            {
                item.TextColor=Color.Black;
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void richboxReceive_TextChanged(object sender, EventArgs e)
        {

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }
    }
}
