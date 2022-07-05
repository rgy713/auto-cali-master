using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace comTest
{
    public class Back5 : Back
    {
        public  byte[] DataBytes { get; set; }
        public override int Length
        {
            get
            {
                return 58;
            }
        }

        public byte codeId
        {
            get
            {
                return DataBytes[0];
            }
        }

        //体重系统
        public RATIO WeightRatio
        {
            get
            {
                RATIO ratio;
                ratio.K = BitConverter.ToSingle(DataBytes, 1);
                ratio.B = BitConverter.ToSingle(DataBytes, 5);
                return ratio;
            }
        }

        //通道类型,未使用

        //通道1系统
        public RATIO Resistance1Ratio
        {
            get
            {
                RATIO ratio;
                ratio.K = BitConverter.ToSingle(DataBytes, 10);
                ratio.B = BitConverter.ToSingle(DataBytes, 14);
                return ratio;
            }
        }
        public RATIO Resistance2Ratio
        {
            get
            {
                RATIO ratio;
                ratio.K = BitConverter.ToSingle(DataBytes,18);
                ratio.B = BitConverter.ToSingle(DataBytes,22);
                return ratio;
            }
        }
        public RATIO Resistance3Ratio
        {
            get
            {
                RATIO ratio;
                ratio.K = BitConverter.ToSingle(DataBytes, 26);
                ratio.B = BitConverter.ToSingle(DataBytes, 30);
                return ratio;
            }
        }
        public RATIO Resistance4Ratio
        {
            get
            {
                RATIO ratio;
                ratio.K = BitConverter.ToSingle(DataBytes, 34);
                ratio.B = BitConverter.ToSingle(DataBytes, 38);
                return ratio;
            }
        }
        public RATIO Resistance5Ratio
        {
            get
            {
                RATIO ratio;
                ratio.K = BitConverter.ToSingle(DataBytes, 42);
                ratio.B = BitConverter.ToSingle(DataBytes, 46);
                return ratio;
            }
        }
        public RATIO Resistance6Ratio
        {
            get
            {
                RATIO ratio;
                ratio.K = BitConverter.ToSingle(DataBytes, 50);
                ratio.B = BitConverter.ToSingle(DataBytes, 54);
                return ratio;
            }
        }
    }
}
