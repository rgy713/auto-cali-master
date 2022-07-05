using Common;
using System;
using System.Windows.Forms;
using System.Collections.Generic;

namespace comTest
{
    class Util
    {
        public static T CreateJaggedArray<T>(object defaultValue, params int[] lengths)
        {
            return (T)InitializeJaggedArray(typeof(T).GetElementType(), 0, lengths, defaultValue);
        }

        public static object InitializeJaggedArray(Type type, int index, int[] lengths, object defaultValue)
        {
            Array array = Array.CreateInstance(type, lengths[index]);
            Type elementType = type.GetElementType();

            if (elementType != null)
            {
                for (int i = 0; i < lengths[index]; i++)
                {
                    array.SetValue(
                    InitializeJaggedArray(elementType, index + 1, lengths, defaultValue), i);
                }
            }
            else
            {
                for (int i = 0; i < lengths[index]; i++)
                {
                    array.SetValue(defaultValue, i);
                }
            }

            return array;
        }


        public static float toFloat(string str)
        {
            uint num = uint.Parse(str, System.Globalization.NumberStyles.AllowHexSpecifier);
            byte[] floatVals = BitConverter.GetBytes(num);
            return BitConverter.ToSingle(floatVals, 0);
        }
        //交换字符串
        public static string exchange(string temp)
        {
            char[] cs = temp.ToCharArray();
            char[] cTemp = new char[] { cs[6], cs[7], cs[4], cs[5], cs[2], cs[3], cs[0], cs[1] };
            return new string(cTemp);
        }
        public static string logBytes(byte[] bytes)
        {
            string logMsg = string.Empty;
            foreach (byte b in bytes)
            {
                logMsg += b.ToString("x2");
                logMsg += "  ";
            }
            return logMsg;
        }

        public static string logListBytes(List<byte> bytes)
        {
            string logMsg = string.Empty;
            foreach (byte b in bytes)
            {
                logMsg += b.ToString("x2");
                logMsg += "  ";
            }
            return logMsg;
        }

        public static void textValue2Ini(string key, TextBox txt)
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

        public static void IniValue2Combox(string key, ComboBox cbx)
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

        public static void IniValue2Text(string key, TextBox txtb)
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

        public static void cbxValue2Ini(string key, ComboBox cbx)
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
        public static byte getCRC(byte[] value, int index, int number)
        {
            int res = 0;
            for (int i = 0; i < number; i++)
            {
                res += value[i + index];
            }
            res = (byte)((res & 0x3F) + 10);
            return (byte)res;
        }
    }
}
