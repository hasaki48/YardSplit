using INFITF;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace YardSplit
{
    internal class Common
    {
        // 隐藏某个部件
        internal static void Hide(INFITF.Application CATIA, object obj)
        {
            Selection selection = CATIA.ActiveDocument.Selection;
            selection.Clear();
            selection.Add((AnyObject)obj);
            CATIA.StartCommand("隐藏/显示");
        }
    }
}

