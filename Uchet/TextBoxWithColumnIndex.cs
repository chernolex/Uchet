using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Uchet
{
    class TextBoxWithColumnIndex
    {
        public TextBox tb;
        public int columnIndex;

        public TextBoxWithColumnIndex()
        {
            tb = new TextBox();
            columnIndex = 0;
        }

        public TextBoxWithColumnIndex(TextBox _tb, int CI)
        {
            tb = _tb;
            columnIndex = CI;
        }
    }
}
