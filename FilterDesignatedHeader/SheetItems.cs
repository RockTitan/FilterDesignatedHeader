using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FilterDesignatedHeader
{
    public class SheetItems
    {
        public string SheetName { get; set; }
        public string HeaderItem { get; set; }
    }

    public class MatchItems
    {
        public string SelectedHeader { get; set; }
        public string MatchItem { get; set; }
        public int MatchIndex { get; set; }
        public string Result { get; set; }
    }
}
