using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace comTest
{
    class Command
    {
        private byte m_header;
        private byte m_second;
        private byte m_footer;
        private int m_startCRC;

        public Command(byte header, byte second, byte footer, int startCRC)
        {
            m_header = header;
            m_second = second;
            m_footer = footer;
            m_startCRC = startCRC;
        }

        public byte[] Make(byte[] values)
        {
            int length = values.Length;
            byte[] command = new byte[length + 6];
            command[0] = m_header;
            command[1] = m_second;
            command[2] = (byte) ((length & 0x3F) + 10);
            command[3] = (byte) (((length >> 6) & 0x3F) + 10);

            Array.Copy(values, 0, command, 4, length);

            command[length + 4] = Util.getCRC(command, m_startCRC, length + 4 - m_startCRC);
            command[length + 5] = m_footer;

            return command;
        }
    }
}
