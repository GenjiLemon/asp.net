using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FileDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            StreamWriter sw = File.CreateText("D://1.doc");
            sw.Write("hahhass");
            sw.Write("dierge");
            
            sw.Close();
        }
    }
}
