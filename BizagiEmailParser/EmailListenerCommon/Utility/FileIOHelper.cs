using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EmailListenerCommon
{
    public static class FileIOHelper
    {
        private static StreamWriter sw;
        public static void WriteToNewFile(string path, string fileNameWithExtension, string Content)
        {
            var fileCompletePath = string.Format("{0}\\{1}", path, fileNameWithExtension);
            if (System.IO.File.Exists(fileCompletePath))
                System.IO.File.Delete(fileCompletePath);
            sw = new System.IO.StreamWriter(System.IO.File.Create(fileCompletePath));
            sw.Close();
            File.WriteAllText(fileCompletePath, Content + Environment.NewLine);
        }

        public static void AppendToFile(string path, string fileNameWithExtension, string Content)
        {
            var fileCompletePath = string.Format("{0}\\{1}", path, fileNameWithExtension);
            File.AppendAllText(fileCompletePath, Content + Environment.NewLine);
        }
    }
}
