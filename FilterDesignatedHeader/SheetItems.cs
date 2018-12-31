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

    public class SheetItemsComparer : IEqualityComparer<SheetItems>
    {
        public bool Equals(SheetItems x, SheetItems y)
        {
            return x.HeaderItem.ToUpper() == y.HeaderItem.ToUpper();
        }
        public int GetHashCode(SheetItems obj)
        {
            return obj.HeaderItem.ToUpper().GetHashCode();
        }
    }

    public class MatchItems
    {
        public string SelectedHeader { get; set; }
        public string MatchItem { get; set; }
        public int MatchIndex { get; set; }
        public string Result { get; set; }
    }
}
