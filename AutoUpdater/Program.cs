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
            App.Run(new Form1());
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
                Console.WriteLine("New release is available");
                var asset = client.Repository.Release.GetAsset("Peshiin", "Skompare", latestRelease.Id);
                Console.WriteLine(asset.Id);

                string downloadUrl = $"https://api.github.com/repos/Peshiin/Skompare/releases/assets/{asset.Id}";
                Uri downloadUri = new Uri(downloadUrl);

                // Download with WebClient
                var webClient = new WebClient();

                // Download the file
                webClient.DownloadFileAsync(downloadUri, "C:/Users/pechm/Desktop/AutoUpdater/File.zip");

            }
           
            Console.WriteLine("Ending");
        }
    }


}
