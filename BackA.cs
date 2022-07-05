using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace comTest
{
    public class BackA : Back
    {
        public  byte[] DataBytes { get; set; }
        public override int Length
        {
            get
            {
                return 42;
            }
        }

        public byte codeId
        {
            get
            {
                return DataBytes[0];
            }
        }

        //D1： 4个字节，float型，体重系数Kcal；
        public float 体重K
        {
            get
            {
                return BitConverter.ToSingle(DataBytes, 2);
            }
        }
        //D2： 4个字节，float型，体重系数Bcal；
        public float 体重B
        {
            get
            {
                return BitConverter.ToSingle(DataBytes, 6);
            }
        }

        //D3： 4个字节，float型，5Khz激励下通道1阻抗系数K；
        public float 阻抗5K1
        {
            get
            {
                return BitConverter.ToSingle(DataBytes, 10);
            }
        }
        //D4： 4个字节，float型，5Khz激励下通道1阻抗系数B；
        public float 阻抗5B1
        {
            get
            {
                return BitConverter.ToSingle(DataBytes, 14);
            }
        }
        //D5： 4个字节，float型，5Khz激励下通道2阻抗系数K；
        public float 阻抗5K2
        {
            get
            {
                return BitConverter.ToSingle(DataBytes, 18);
            }
        }
        //D6： 4个字节，float型，5Khz激励下通道2阻抗系数B；
        public float 阻抗5B2
        {
            get
            {
                return BitConverter.ToSingle(DataBytes, 22);
            }
        }
        //D7： 4个字节，float型，250Khz激励下通道1阻抗系数K；
        public float 阻抗250K1
        {
            get
            {
                return BitConverter.ToSingle(DataBytes, 26);
            }
        }
        //D8： 4个字节，float型，250Khz激励下通道1阻抗系数B；
        public float 阻抗250B1
        {
            get
            {
                return BitConverter.ToSingle(DataBytes, 30);
            }
        }
        //D9： 4个字节，float型，250Khz激励下通道2阻抗系数K；
        public float 阻抗250K2
        {
            get
            {
                return BitConverter.ToSingle(DataBytes, 34);
            }
        }
        //D10： 4个字节，float型，250Khz激励下通道2阻抗系数B；
        public float 阻抗250B2
        {
            get
            {
                return BitConverter.ToSingle(DataBytes, 38);
            }
        }

        ////频段号
        //public byte 频段号
        //{
        //    get
        //    {
        //        return DataBytes[1];
        //    }
        //}
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
