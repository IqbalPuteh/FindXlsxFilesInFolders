using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FindXlsxFilesInFolders
{
    public class FindXlsxFilesInFolders
    {
        static readonly string fileFilter = "*.xlsx";
        static bool blnShowUnaccessFolder = false;
        static bool blnShowFileSize = false;

        static void Main()
        {
            try
            {
                //Show the apps Welcome Header
                ShowAppHeaderWelcome();

                // Start with all drives in computer.
                string[] drives = System.Environment.GetLogicalDrives();

                foreach (string dr in drives)
                {
                    System.IO.DriveInfo di = new System.IO.DriveInfo(dr);

                    // Skip the drive if it is not ready to be read.
                    if (!di.IsReady)
                    {
                        Console.WriteLine("The drive {0} could not be read", di.Name);
                        continue;
                    }
                    System.IO.DirectoryInfo rootDir = di.RootDirectory;
                    MyGetDirectories(rootDir.ToString());
                }
           
                Console.WriteLine("Press any key..");
                Console.ReadKey();
            }
            catch (Exception e)
            {
                // Catch  exceptions 
                Console.WriteLine(e.Message.ToString());
            }
        }

        public static void MyGetDirectories(string folderPath)
        {
            string[] directories = Directory.GetDirectories(folderPath);

            foreach (string directory in directories)
            {
                try
                {
                    MyGetDirectories(directory);

                    //Console.WriteLine($"{directory}");
                }
                catch (UnauthorizedAccessException)
                {
                    if (blnShowUnaccessFolder)
                    {
                        Console.WriteLine($"Can't access directory {directory}, skip.");
                    }
                }
                catch (Exception e)
                {
                    // Catch remaining exceptions 
                    Console.WriteLine(e.Message.ToString());
                }
            }

            string[] files = Directory.GetFiles(folderPath, fileFilter);

            foreach (string file in files)
            {
                try
                {
                    FileInfo fileInfo = new FileInfo(file);

                    if (blnShowFileSize)
                    {
                        Console.WriteLine($"File Name = {file} - File Size = {fileInfo.Length}");
                    } 
                    else
                    {
                        Console.WriteLine($"File Name = {file}");
                    }
                }
                catch (UnauthorizedAccessException)
                {
                    Console.WriteLine($"Cannot access file information {file}.");
                }
                catch (Exception e)
                {
                    // Catch remaining exceptions 
                    Console.WriteLine(e.Message.ToString());
                }
            }
        }

        public static void ShowAppHeaderWelcome()
        {
            Console.WriteLine("**********************************************************************");
            Console.WriteLine("*                                                                    *");
            Console.WriteLine("*        App purpose : Find all *.xlxs file in all drives            *");
            Console.WriteLine("*        Created by  : Iqbal Reza Puteh                              *");
            Console.WriteLine("*        Created Date: September 23rd, 2022                          *");
            Console.WriteLine("*        Version    ; 1.0.0                                          *");
            Console.WriteLine("*                                                                    *");
            Console.WriteLine("**********************************************************************");
            Console.WriteLine();
            Console.Write("Do you want to show un-accesible folder name(s) (Y/N)?");

            if (Console.ReadLine() == "Y")
            {
                blnShowUnaccessFolder = true;
            }
            Console.Write("Do you want to show file(s) size (Y/N)?");
            if (Console.ReadLine() == "Y")
            {
                blnShowFileSize = true;
            }
        }
    }
}
