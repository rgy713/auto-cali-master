using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace comTest
{
    public class BackC : Back
    {
        public  byte[] DataBytes { get; set; }
        public override int Length
        {
            set;
            get;
        }

        public byte codeId
        {
            get
            {
                return DataBytes[0];
            }
        }

        //频段号
        public byte 数据类型
        {
            get
            {
                return DataBytes[1];
            }
        }

        public string GetDataType()
        {
            switch (数据类型)
            {
                case 3:
                    return "byte";
                case 1:
                    return "ushort";
                case 2:
                    return "short";
                case 0:
                    return "float";
            }
            return string.Empty;
        }

        private string readByteStr()
        {
            string temp = string.Empty;
            byte[] datas = new byte[Length - 2];
            Array.Copy(DataBytes, 0, datas, 2, datas.Length);
            foreach(byte item in datas)
            {
                temp += item.ToString("x2");
                temp += ", ";
            }
            return temp;
        }

        private string readUshortStr()
        {
            string temp = string.Empty;
            byte[] datas = new byte[Length - 2];
            Array.Copy(DataBytes, 0, datas, 2, datas.Length);

            for (int i = 0; i < datas.Length / 2; i++ )
            {
                ushort uValue = BitConverter.ToUInt16(datas, 2*i);
                temp += uValue.ToString("x2");
                temp += ", ";
            }
            return temp;
        }

        private string readShortStr()
        {
            string temp = string.Empty;
            byte[] datas = new byte[Length - 2];
            Array.Copy(DataBytes, 0, datas, 2, datas.Length);

            for (int i = 0; i < datas.Length / 2; i++)
            {
                short uValue = BitConverter.ToInt16(datas, 2 * i);
                temp += uValue.ToString("x2");
                temp += ", ";
            }
            return temp;
        }
        private string readFloatStr()
        {
            string temp = string.Empty;
            byte[] datas = new byte[Length - 2];
            Array.Copy(DataBytes, 0, datas, 2, datas.Length);

            for (int i = 0; i < datas.Length / 2; i++)
            {
                float uValue = BitConverter.ToSingle(datas, 4 * i);
                temp += uValue.ToString("x2");
                temp += ", ";
            }
            return temp;
        }

        //读取string
        public string GetValueStr()
        {
            string tempStr = string.Empty;
            switch (数据类型)
            {
                case 3:
                    tempStr=readByteStr();
                    break;
                case 1:
                    tempStr = readUshortStr();
                    break;
                case 2:
                    tempStr = readShortStr();
                    break;
                case 0:
                    tempStr = readFloatStr();
                    break;
            }
            return tempStr;
        }


        ////体重传感器数值
        //public float 体重
        //{
        //    get
        //    {
        //        return BitConverter.ToSingle(DataBytes, 2);
        //    }
        //}
        ////通道类型,未使用 1

        ////通道1系统
        //public float 阻抗1
        //{
        //    get
        //    {
        //        return BitConverter.ToSingle(DataBytes, 7);
        //    }
        //}
        //public float 阻抗2
        //{
        //    get
        //    {
        //        return BitConverter.ToSingle(DataBytes, 11);
        //    }
        //}
        //public float 阻抗3
        //{
        //    get
        //    {
        //        return BitConverter.ToSingle(DataBytes, 15);
        //    }
        //}
        //public float 阻抗4
        //{
        //    get
        //    {
        //        return BitConverter.ToSingle(DataBytes, 19);
        //    }
        //}
        //public float 阻抗5
        //{
        //    get
        //    {
        //        return BitConverter.ToSingle(DataBytes, 23);
        //    }
        //}
        //public float 阻抗6
        //{
        //    get
        //    {
        //        return BitConverter.ToSingle(DataBytes, 27);
        //    }
        //}
    }
}
