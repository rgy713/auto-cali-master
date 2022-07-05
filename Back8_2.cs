using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace comTest
{
    public class Back8_2 : Back
    {
        public  byte[] DataBytes { get; set; }
        public override int Length
        {
            get
            {
                return 15;
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
     
    }
}
