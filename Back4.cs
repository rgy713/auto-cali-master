using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace comTest
{

    public class Back4 : Back
    {
        public  byte[] DataBytes { get; set; }
        public override int Length
        {
            get
            {
                return 17;
            }
        }

        public byte codeId
        {
            get
            {
                return DataBytes[0];
 
            }
        }
        public byte 频段号
        {
            get
            {
                return DataBytes[1];

            }
        }
        public UInt16 体重
        {
            get
            {
                return BitConverter.ToUInt16(DataBytes, 2);
            }
        }
        public byte 通道类型
        {
            get
            {
                return DataBytes[4];

            }
        }
        public UInt16 阻抗1
        {
            get
            {
               return BitConverter.ToUInt16(DataBytes, 5);
            }
        }
        public UInt16 阻抗2
        {
            get
            {
                return BitConverter.ToUInt16(DataBytes, 7);

            }
        }
        public UInt16 阻抗3
        {
            get
            {
                return BitConverter.ToUInt16(DataBytes, 9);

            }
        }
        public UInt16 阻抗4
        {
            get
            {
                return BitConverter.ToUInt16(DataBytes, 11);
            }
        }
        public UInt16 阻抗5
        {
            get
            {
                return BitConverter.ToUInt16(DataBytes, 13);
            }
        }
        public UInt16 阻抗6
        {
            get
            {
                return BitConverter.ToUInt16(DataBytes,15);
            }
        }
    }
}
