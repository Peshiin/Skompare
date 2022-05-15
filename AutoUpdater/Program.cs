using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Octokit;
using System.Reflection;
using Skompare;
using System.Net;
using System.Net.Http;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace AutoUpdater
{
    internal static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            AutoUpdaterClass demo = new AutoUpdaterClass();
            demo.Update();
        }
    }

    public class AutoUpdaterClass
    {

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        public void Update()
        {
            // Sets client for Github access
            var client = new GitHubClient(new ProductHeaderValue("SkompareUpdate"));

            // Gets the latest release of the application
            var latestRelease = client.Repository.Release.GetLatest("Peshiin", "Skompare").Result;
            Console.WriteLine(latestRelease.TagName);

            //Gets location of current assembly executable - AutoUpdater.exe
            string assemblyPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
            //Gets path to current assembly directory
            string dirPath = new FileInfo(assemblyPath).DirectoryName;
            //Gets version info of Skompare.exe
            var versionInfo = FileVersionInfo.GetVersionInfo(dirPath+"\\Skompare.exe");
            //Extracts version number from version info
            string version = versionInfo.FileVersion;

            if (latestRelease.TagName != version)
            {
                //Shows the dialog of different versions
                DialogResult dialogResult = MessageBox.Show("Chcete nainstalovat poslední verzi aplikace?" +
                                                            Environment.NewLine +
                                                            version +"->"+latestRelease.TagName,
                                                           "Close",
                                                           MessageBoxButtons.YesNo,
                                                           MessageBoxIcon.Question);
                
                if (dialogResult == DialogResult.Yes)
                {
                    // Download with WebClient
                    var webClient = new WebClient();
                    webClient.Headers.Add(HttpRequestHeader.UserAgent, "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/70.0.3538.77 Safari/537.36");

                    //Gets URL for downloading the asset
                    var downloadUrl = string.Format("https://github.com/Peshiin/Skompare/releases/download/{0}/release.zip", latestRelease.TagName);
                    //Creates file in the memory to store the asset
                    byte[] fileInMemory = webClient.DownloadData(downloadUrl);
                    //Saves the downloaded file to a directory
                    File.WriteAllBytes(Path.Combine(dirPath, "newRelease.zip"), fileInMemory);

                    using (var archive = ZipFile.OpenRead(dirPath + "\\newRelease.zip"))
                    {
                        //Unzip all the files to a set directory
                        foreach (ZipArchiveEntry entry in archive.Entries) 
                        {
                            //Unzip a file in archive, true is for allowing overwriting existing files
                            entry.ExtractToFile(Path.Combine(dirPath, entry.FullName), true);
                        }
                        //Disposes archive to allow deleting
                        archive.Dispose();
                        //Deletes archive
                        File.Delete(dirPath + "\\newRelease.zip");
                    }
                }
            }
            //Starts the updated application
            Process process = Process.Start(dirPath + "\\Skompare.exe");
            //Gets window handle for the started process to set it on foreground
            IntPtr processHandle = process.MainWindowHandle;
            //Sets the app window on foreground
            SetForegroundWindow(processHandle);
        }
    }


}
