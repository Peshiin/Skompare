using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using App = System.Windows.Forms.Application;
using Octokit;
using System.Reflection;
using Skompare;
using System.Net;
using System.Net.Http;
using System.IO;
using System.Text;

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
            App.EnableVisualStyles();
            App.SetCompatibleTextRenderingDefault(false);
            //App.Run(new Form1());
            DemoClass demo = new DemoClass();
            demo.Demo();
        }


    }

    public class DemoClass
    {
        public async void Demo()
        {
            // Sets client for Github access
            var client = new GitHubClient(new ProductHeaderValue("SkompareUpdate"));

            // Gets the latest release of the application
            var latestRelease = client.Repository.Release.GetLatest("Peshiin", "Skompare").Result;
            Console.WriteLine(latestRelease.TagName);

            // Gets the version of current assembly
            string version = Assembly.GetExecutingAssembly().GetName().Version.ToString();
            Console.WriteLine(version);

            if (latestRelease.TagName != version)
            {

                // Download with WebClient
                var webClient = new WebClient();
                webClient.Headers.Add(HttpRequestHeader.UserAgent, "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/70.0.3538.77 Safari/537.36");

                // Download the file
                var downloadLocation = @"C:\Users\n5ver\Desktop\skomparedown";

                //File.WriteAllBytes(Path.Combine(downloadLocation, "test.zip"), inMemoryFile);
                var downloadUrl = string.Format("https://github.com/Peshiin/Skompare/releases/download/{0}/release.zip", latestRelease.TagName);
                byte[] fileInMemory = webClient.DownloadData(downloadUrl);

                File.WriteAllBytes(Path.Combine(downloadLocation, "test.zip"), fileInMemory);
            }
           
            Console.WriteLine("Ending");
        }
    }


}
