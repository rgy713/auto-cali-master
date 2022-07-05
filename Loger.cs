using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace Common
{
    public delegate void DelegateLog(string msg);
    public class Loger
    {

        public static event DelegateLog LogRecived;
        public static event DelegateLog LogSend;

        public static  void send(string msg)
        {
            string stime = DateTime.Now.ToString("HH:mm:ss.fff") + ":  ";

            if (LogSend != null) LogSend(stime+msg);
        }
        public static void recv(string msg)
        {
            string stime = DateTime.Now.ToString("HH:mm:ss.fff") + ":  ";

            if (LogRecived != null) LogRecived(stime+msg);
        }
    }
}
