using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace SharePointGraphClientSecret
{
    internal class LinkToDocument
    {
        public static async Task Create(GraphServiceClient graphClient, Site site)
        {
            var devCRMSite = graphClient.Sites[site.Id];



            var taskDrive1 = devCRMSite.Drives
                .Request().GetAsync().GetAwaiter().GetResult();
            var test = taskDrive1[1];

            var testGA = devCRMSite.Drives[test.Id];
            var testResult = testGA.Request().GetAsync().GetAwaiter().GetResult();

            var folders = await testGA
             .Request()
    .Expand("Fields")
    .Filter("Fields/Archived eq 'false'");

            //var letterDrive = await devCRMSite.Drives
            //  .Request()
            //.GetAsync((requestConfiguration) =>
            // {
            //     requestConfiguration.QueryParameters.Filter = "fields/Title eq 'Contoso Home'";
            //     requestConfiguration.Headers.Add("Prefer", "allowthrottleablequeries");
            // });

 //           var letterDrive = await devCRMSite.Drives
 //   .Request()
 //.Filter("Name eq 'letter'")
 //.GetAsync()
 //;


            Console.WriteLine();
            //var letterDrive2 = await devCRMSite.Drives
            //    .Request()
            //.GetAsync((requestConfiguration) =>
            //{
            //    requestConfiguration.QueryParameters.Filter = "Name eq 'Letter'";
            //})
            //.GetAwaiter().GetResult();


            var taskDrive1 = devCRMSite.Drives
                .Request().GetAsync().GetAwaiter().GetResult();
            var test = taskDrive1[1];

            var testGA = devCRMSite.Drives[test.Id];
            var testResult = testGA.Request().GetAsync().GetAwaiter().GetResult();

            //var folder = graphClient.Sites[site.Id].Drives[test.Id].Items.Request()
            //    .Filter("Name eq 'privContact'").GetAsync()
            //    .GetAwaiter().GetResult();

            //.GetAsync()
            //.GetAwaiter().GetResult();
            //CreateFileInFolder(graphClient, site, drive1[0], new MemoryStream(new byte[] { 40,40,40,40,41,41,41,41}));
        }

        private static async void CreateFileInFolder(GraphServiceClient graphClient, Site site, DriveItem newFolder, MemoryStream stream)
        {
            DriveItem file = await graphClient
              .Sites[site.Id]
              .Drives["Letter"]
              .Items[newFolder.Id]
              .ItemWithPath("letter/" + newFolder.Name)
              .Content
                        .Request()
              .PutAsync<DriveItem>(stream);
        }

        private void GetTheFilesWithinAFolder(GraphServiceClient graphClient, Site site, string newFolderName)
        {
            // This gets the folder within a speciifc Document library
            var taskSearch = graphClient.Sites[site.Id].GetByPath($"/Letter/newFolderName/").Request().GetAsync();
            var searchResult = taskSearch.GetAwaiter().GetResult();
        }

        private static void GetNewlyCreatedFolderObject(GraphServiceClient graphClient, Site site, string newFolderName)
        {
            // This gets the folder within a speciifc Document library
            var taskSearch = graphClient.Sites[site.Id].Lists["Letter"].Drive.Root.Children[newFolderName].Request().GetAsync();
            var searchResult = taskSearch.GetAwaiter().GetResult();
        }

        private static void CreateANewFolder(GraphServiceClient graphClient, Site site, DriveItem newFolder)
        {
            // This creates a new folder
            var taskDriveLetter = graphClient.Sites[site.Id].Lists["Letter"].Drive.Root.Children.Request().AddAsync(newFolder);
            var driveLetter = taskDriveLetter.GetAwaiter().GetResult();
        }
    }
}
