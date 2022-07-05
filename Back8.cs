using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace comTest
{
    public class Back8 : Back
    {
        public  byte[] DataBytes { get; set; }
        public override int Length
        {
            get
            {
                return 31;
            }
        }

        public byte codeId
        {
            get
            {
                return DataBytes[0];
            }
        }

        //频段号
        public byte 频段号
        {
            get
            {
                return DataBytes[1];
            }
        }
        //体重传感器数值
        public float 体重
        {
            get
            {
                return BitConverter.ToSingle(DataBytes, 2);
            }
        }
        //通道类型,未使用 1

        //通道1系统
        public float 阻抗1
        {
            get
            {
                return BitConverter.ToSingle(DataBytes, 7);
            }
        }
        public float 阻抗2
        {
            get
            {
                return BitConverter.ToSingle(DataBytes, 11);
            }
        }
        public float 阻抗3
        {
            get
            {
                return BitConverter.ToSingle(DataBytes, 15);
            }
        }
        public float 阻抗4
        {
            get
            {
                return BitConverter.ToSingle(DataBytes, 19);
            }
        }
        public float 阻抗5
        {
            get
            {
                return BitConverter.ToSingle(DataBytes, 23);
            }
        }
        public float 阻抗6
        {
            get
            {
                return BitConverter.ToSingle(DataBytes, 27);
            }
        }
    }
}
