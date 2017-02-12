using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Shell32;

namespace ImportAudio
{
    class ImportAudio
    {
        [STAThread]
        static void Main(string[] args){

            //User input for root folder path
            Console.WriteLine("Browse to the root music folder");
            string folderPath = "";
            FolderBrowserDialog fbd = new FolderBrowserDialog();

            //Await user input
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                folderPath = fbd.SelectedPath;
                Shell32.Folder musicFolder = GetShell32Folder(fbd.SelectedPath);
                Console.WriteLine("Name, Size, Item Type, Owner, Album, Year, Genre");
                GenerateLib(musicFolder, folderPath);
            }
        }

        //Create Shell32 Folder
        //NOTE : Shell32.Shell shell = new Shell32.Shell(); no longer works in windows 8, 10
        static private Shell32.Folder GetShell32Folder(string folderPath){
            Type shellAppType = Type.GetTypeFromProgID("Shell.Application");
            Object shell = Activator.CreateInstance(shellAppType);
            return (Shell32.Folder)shellAppType.InvokeMember("NameSpace",
            System.Reflection.BindingFlags.InvokeMethod, null, shell, new object[] { folderPath });
        }

        //Generates list of songs within the passed directory
        static private void GenerateLib(Shell32.Folder dir, String path){

            foreach (Shell32.FolderItem2 item in dir.Items())
            {
                Console.WriteLine("{0},{1},{2},{3},{4},{5},{6}", dir.GetDetailsOf(item, 0), dir.GetDetailsOf(item, 1), dir.GetDetailsOf(item, 2), dir.GetDetailsOf(item, 10), dir.GetDetailsOf(item, 14), dir.GetDetailsOf(item, 15), dir.GetDetailsOf(item, 16));
                //if there is a folder check inside the folder
                if (dir.GetDetailsOf(item, 2) == "File folder")
                {
                    //Go to the folder
                    String folderPath = path + "\\" + dir.GetDetailsOf(item, 0);
                    Console.WriteLine(folderPath);
                    Shell32.Folder weMustGoDeeper = GetShell32Folder(folderPath);

                    //Generate the list of songs in the folderw
                    GenerateLib(weMustGoDeeper, folderPath);
                }
            }
        }
    }
}
