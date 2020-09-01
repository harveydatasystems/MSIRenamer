using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using WindowsInstaller;
using System.IO;

namespace MSIVerRenamer
{
    class Program
    {
        static void Main(string[] args)
        {
            string inputFile;           

            if (args.Length == 0)
            {
                Console.WriteLine("Enter MSI file:");
                inputFile = Console.ReadLine();
            }
            else            
                inputFile = args[0];            

            try
            {
                Console.WriteLine($"Processing Input: {inputFile}");
                if (!System.IO.File.Exists(inputFile))
                    return;

                System.IO.FileInfo fi = new FileInfo(inputFile);

                if (fi.Extension != ".msi")
                    return;

                // Read the MSI property
                //string productName = GetMsiProperty(inputFile, "ProductName");
                string version = GetMsiProperty(inputFile, "ProductVersion");                                
                string directory = fi.DirectoryName;
                string filename = fi.Name.TrimEnd(fi.Extension.ToCharArray());
                string destfile = $"{directory}\\{filename}.{version}{fi.Extension}";
                File.Copy(inputFile, destfile);
                //File.Delete(inputFile);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        static string GetMsiProperty(string msiFile, string property)
        {
            string retVal = string.Empty;

            // Create an Installer instance  
            Type classType = Type.GetTypeFromProgID("WindowsInstaller.Installer");
            Object installerObj = Activator.CreateInstance(classType);
            Installer installer = installerObj as Installer;

            // Open the msi file for reading  
            // 0 - Read, 1 - Read/Write  
            Database database = installer.OpenDatabase(msiFile, 0);

            // Fetch the requested property  
            string sql = String.Format(
                "SELECT Value FROM Property WHERE Property='{0}'", property);
            View view = database.OpenView(sql);
            view.Execute(null);

            // Read in the fetched record  
            Record record = view.Fetch();
            if (record != null)
            {
                retVal = record.get_StringData(1);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(record);
            }
            view.Close();
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(view);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(database);

            return retVal;
        }
    }
}
