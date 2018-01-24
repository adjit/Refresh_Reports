using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Refresh_Reports
{
    class Program
    {
        public const string xmlFilePath = @"File_Paths.xml";
        public const int PARAM = 0;
        public const int FP = 1;

        static void Main(string[] args)
        {
            List<string> filesToRefresh = new List<string>();
            XDocument xmlFile;
            FileInfo xmlFileInfo = new FileInfo(xmlFilePath);
            xmlFileInfo.Directory.Create();

            if(!File.Exists(xmlFilePath))
            {
                xmlFile = new XDocument(new XElement("root", new XElement("FilePaths")));
                xmlFile.Save(xmlFilePath);
            }
            else
            {
                xmlFile = XDocument.Load(xmlFilePath);

                foreach (XElement elem in xmlFile.Root.Element("FilePaths").Elements())
                {
                    filesToRefresh.Add(elem.Value);
                }
            }

            if (args.Length == 0)
            {
                RunSchedule();
            }
            else if (args[PARAM] == "-d")
            {
                for (int i = 0; i < filesToRefresh.Count; i++)
                {
                    Console.WriteLine("[{0}] - {1}", i, filesToRefresh[i]);
                }
                Console.Write("Type index you wish to delete: ");
                Console.Out.Flush();

                int index = Convert.ToInt32(Console.ReadLine());

                deleteFilePath(filesToRefresh[index], xmlFile);
                filesToRefresh.RemoveAt(index);
            }
            else if (args[PARAM] == "-a")
            {
                addFilePath(args[FP], xmlFile);
                filesToRefresh.Add(args[FP]);
            }
            else if (args.Length == 1 && args[PARAM] != "-d")
            {
                RunSchedule(Convert.ToString(args[PARAM]));
            }
        }

        public static void deleteFilePath(string filePath, XDocument xmlFile)
        {
            filePath = filePath.Trim();
            xmlFile.Descendants("FilePaths")
                .Elements("filePath")
                .Where(x => x.Value == filePath)
                .Remove();
            xmlFile.Save(xmlFilePath);
            Console.WriteLine("~~~ {0} - Deleted ~~~", filePath);
        }

        public static void addFilePath(string filePath, XDocument xmlFile)
        {
            xmlFile.Root.Element("FilePaths").Add(new XElement("filePath", filePath));
            xmlFile.Save(xmlFilePath);
            Console.WriteLine("~~~ {0} - Added ~~~", filePath);
        }

        public static void RunSchedule()
        {
            Console.Write("Can you read me? Y/N - ");
            Console.Out.Flush();
            string response = Convert.ToString(Console.ReadLine());
            if (response.ToUpper() == "Y")
            {
                Console.WriteLine("Working!");
            }
            Console.Write("Exit? Y/N - ");
            Console.Out.Flush();

            response = Convert.ToString(Console.ReadLine());
            
        }

        public static void RunSchedule(string filePath)
        {
            Excel.Application exApp = new Excel.Application();

        }
    }
}
