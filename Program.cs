using HybridShapeTypeLib;
using INFITF;
using MECMOD;
using PARTITF;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace YardSplit
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Application CATIA = (INFITF.Application)Marshal.GetActiveObject("Catia.Application");

            string task = "HorizontalSplit";

            switch (task)
            {
                case "HorizontalSplit":
                    HorizontalSplit.HorizontalSplitProcessControl(CATIA);
                    break;
                case "MeasureVolume":
                    MeasureVolume.MeasureVolumeProcessControl(CATIA);
                    break;
            }

        }


    }
}
