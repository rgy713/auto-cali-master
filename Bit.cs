using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Common.Do
{
    public class Bit
    {
        //index从0开始  
        //获取取第index位  
        public static int GetBit(byte b, int index) { return ((b & (1 << index)) > 0) ? 1 : 0; }
        //将第index位设为1  
        public static byte SetBit(byte b, int index) { return (byte)(b | (1 << index)); }
        //将第index位设为0   
        public static byte ClearBit(byte b, int index) { return (byte)(b & (byte.MaxValue - (1 << index))); }
        //将第index位取反   
        public static byte ReverseBit(byte b, int index) { return (byte)(b ^ (byte)(1 << index)); }
    }
}
