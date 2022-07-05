using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace comTest
{
    public class Back7 : Back
    {
        public  byte[] DataBytes { get; set; }
        public override int Length
        {
            get
            {
                return 46;
            }
        }
        public byte codeId
        {
            get
            {
                return DataBytes[0];
            }
        }

        //D1：数据包号：1字节，0~255循环递增，以示数据的连续性。
        public byte 数据包号
        {
            get
            {
                return DataBytes[1];
            }
        }
        //D2：用户编号：8字节字符串
        public string 用户编号
        {
            get
            {
                return BitConverter.ToString(DataBytes, 2, 8);
            }
        }
        //D3：体重值：4字节，float型，单位为kg
        public float 体重值
        {
            get
            {
                return BitConverter.ToSingle(DataBytes, 10);
            }
        }

        //D4：脂肪重：4字节，float型，单位为kg
        public float 脂肪重
        {
            get
            {
                return BitConverter.ToSingle(DataBytes, 14);
            }
        }
        //D5：脂肪率：4字节，float型，单位为%（如56.3代表56.3%）
        public float 脂肪率
        {
            get
            {
                return BitConverter.ToSingle(DataBytes, 18);
            }
        }
        //D6：总体水：4字节，float型，单位为kg
        public float 总体水
        {
            get
            {
                return BitConverter.ToSingle(DataBytes, 22);
            }
        }
        //D7：细胞外液：4字节，float型，单位为kg
        public float 细胞外液
        {
            get
            {
                return BitConverter.ToSingle(DataBytes, 26);
            }
        }

        //D8：肌肉率：4字节，float型，单位为%（如56.3代表56.3%）
        public float 肌肉率
        {
            get
            {
                return BitConverter.ToSingle(DataBytes, 30);
            }
        }

        //D9：蛋白质：4字节，float型，单位为kg
        public float 蛋白质
        {
            get
            {
                return BitConverter.ToSingle(DataBytes, 34);
            }
        }
        //D10：矿物质：4字节，float型，单位为kg
        public float 矿物质
        {
            get
            {
                return BitConverter.ToSingle(DataBytes, 38);
            }
        }
        //D11：基础代谢率：4字节，float型，单位为kcal/d
        public float 基础代谢率
        {
            get
            {
                return BitConverter.ToSingle(DataBytes, 42);
            }
        }
    }
}
