using System;
using System.IO;
using System.IO.Compression;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Text;

namespace h2Print
{
    class Program
    {
        static void Main(string[] args)
        {

            // put driver in temp file
            var archivePath = System.IO.Path.GetTempPath() + "h2driver.zip";
            var folderPath = System.IO.Path.GetTempPath() + "h2driver";
            // clean up artifacts if exist
            try
            {
                File.Delete(archivePath);
                Directory.Delete(folderPath, true);
            }
            catch
            {

            }



            File.WriteAllBytes(archivePath, h2printer2.Properties.Resources.driver);
            ZipFile.ExtractToDirectory(archivePath, folderPath);
            String driverPath = folderPath + "\\driver\\KOAXCJ__.inf";


            var unwantedChars = new Char[] { '\uFEFF', '\u200B' };
            var scriptString = Encoding.UTF8.GetString(h2printer2.Properties.Resources.installH2).Trim(unwantedChars);

            var setDriverPath = String.Format("$DRIVERLOCATION='{0}'", driverPath);

            Console.WriteLine("Adding H2 Printer directly, please wait... (this window will close automatically)");

            Console.WriteLine("Would you like to remove printserver connection (WPMC6021STH2)? This can help if the print screen FREEZES. (y/n)");
            var deleteOldPrinter = Console.ReadLine().ToLower()[0].ToString();

            Console.WriteLine("Got it. Please sit tight, this may take a minute or so :)");

            PowerShell ps = PowerShell.Create();

            //permit script execution
            ps.AddCommand("Set-ExecutionPolicy")
              .AddParameter("Scope", "Process")
              .AddParameter("ExecutionPolicy", "Bypass")
              .Invoke();



            //run embedded script to install printer


            ps
                .AddScript(scriptString)
                .AddParameter("DRIVERLOCATION", driverPath)
                .AddParameter("DELETEOLD", deleteOldPrinter);


            var result = ps.Invoke();
            //foreach (var line in result)
            //{
            //    Console.WriteLine(line);
            //}

            Console.WriteLine("Alright! Should have been successful. Use \'H2 Direct\' to print.");

            Console.ReadLine();

        }

    }
}
