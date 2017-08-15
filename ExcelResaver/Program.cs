using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace ExcelResaver
{
    class Program
    {
        static void Main(string[] args)
        {
            Controler controller = new Controler();
            string[] goodArgs = Helper.SplitBySpaceBar(args[0]);
            controller.ResaveFilesInFolder(goodArgs);
        }
    }
}
