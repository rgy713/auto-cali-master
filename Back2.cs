using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace comTest
{
    public class Back2 : Back
    {
        public  byte[] DataBytes { get; set; }

        public override int Length
        {
            get
            {
                return 1;
            }
        }

        public byte codeId
        {
            get
            {
                return DataBytes[0];

            }
        }
    }
}
