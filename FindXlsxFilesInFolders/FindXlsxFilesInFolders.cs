using System;
using System.IO;
using System.Reflection;

namespace FindXlsxFilesInFolders
{
    public class FindXlsxFilesInFolders
    {
        static readonly string fileFilter = "*.xlsx";
        static bool blnShowUnaccessFolder = false;

        static void Main()
        {
            try
            {
                //Show the apps welcome header
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
                // Catch exceptions 
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

                }
                catch (UnauthorizedAccessException)
                {
                    //Catch UnauthorizedAccessException   
                    if (blnShowUnaccessFolder)
                    {
                        Console.WriteLine($"Can't access directory {directory}, skip.");
                    }
                }
                catch (Exception e)
                {
                    //Show other exception error than UnauthorizedAccessException if blnShowUnaccessFolder flag is true  
                    if (blnShowUnaccessFolder)
                    {
                        Console.WriteLine(e.Message.ToString());
                    }
                    
                }
            }

            string[] files = Directory.GetFiles(folderPath, fileFilter);

            foreach (string file in files)
            {
                try
                {
                    FileInfo fileInfo = new FileInfo(file);

                    Console.WriteLine($"File Name = {file} - File Size = {fileInfo.Length}");
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
            try
            {
                Console.WriteLine("**********************************************************************");
                Console.WriteLine("*                                                                    *");
                Console.WriteLine("*        App purpose : Find all *.xlxs file in all drives            *");
                Console.WriteLine("*        Created by  : Iqbal Reza Puteh                              *");
                Console.WriteLine("*        Created Date: September 23rd, 2022                          *");
                Console.WriteLine("*        Version     : " + Assembly.GetExecutingAssembly().GetName().Version + "                                       *");
                Console.WriteLine("*                                                                    *");
                Console.WriteLine("**********************************************************************");
                Console.WriteLine();
                Console.Write("Do you want to show un-accesible folder name(s) (y/n)?");
                ConsoleKeyInfo x = Console.ReadKey();
                if (x.KeyChar == 'y' || x.KeyChar == 'Y')
                {
                    blnShowUnaccessFolder = true;
                }
                //Move to new line
                Console.WriteLine();
                Console.WriteLine();
            }
            catch (Exception e)
            {
                //Throw exception to main function
                throw e;
            }
        }
    }
}
