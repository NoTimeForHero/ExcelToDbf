using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UnitTests
{
    class TestLibrary
    {

        public static bool safeDelete(string filename)
        {
            try
            {
                if (File.Exists(filename)) File.Delete(filename);
                return true;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Can't delete filename '{filename}'!");
                Console.Error.WriteLine("Error: " + ex.Message);
                return false;
            }
        }

        public static string getTempFilename(string extension)
        {
            var temp = Path.GetTempPath();
            var rndName = Path.GetRandomFileName();
            var outName = Path.ChangeExtension(rndName, extension);
            return Path.Combine(temp, outName);
        }

    }
}
