//// See https://aka.ms/new-console-template for more information
using Azure.Core;
using Microsoft.Graph;
using SharePointGraphClientSecret;

Console.WriteLine("Hello, World!");


var graphClient = (new GraphCServicelientHelper()).GetGraphClient();

//var me = await graphClient.Me.Request().GetAsync();

var task = graphClient.Sites["kamalacon.sharepoint.com:/sites/CRMDEV/"].Request().GetAsync();
//var task = graphClient.Sites["root"].SiteWithPath("/sites/CRMDEV").Request().GetAsync();
var site = task.GetAwaiter().GetResult();

//var taskDrive = graphClient.Sites[site.Id].Drives["Letter"].Request().GetAsync();
var taskDrive = graphClient.Sites[site.Id].Lists["Letter"].Request().GetAsync();
var drive = taskDrive.GetAwaiter().GetResult();

var newFolderName = "newFolder" + DateTime.Now.ToString("yyyyMMddhhmmss");

var newFolder = new DriveItem
{
    Name = newFolderName,
    Folder = new Folder()
};

/*
//  This throws an Exception when there is no folder found under the specific Document Library
var taskSearchBeforeCreate = graphClient.Sites[site.Id].Lists["Letter"].Drive.Root.Children[newFolderName].Request().GetAsync();
var searchResultBeforeCreate = taskSearchBeforeCreate.GetAwaiter().GetResult();
*/

/*ERROR THROWN WHEN THE FOLDER DOES NOT EXISTS

Microsoft.Graph.ServiceException
  HResult=0x80131500
  Message=Code: itemNotFound
Message: The resource could not be found.
Inner error:
	AdditionalData:
	date: 2023-01-08T12:30:20
	request-id: 2b87a77e-c50b-4797-97cf-66dd0d0bcf3f
	client-request-id: 2b87a77e-c50b-4797-97cf-66dd0d0bcf3f
ClientRequestId: 2b87a77e-c50b-4797-97cf-66dd0d0bcf3f

  Source=Microsoft.Graph.Core
  StackTrace:
   at Microsoft.Graph.HttpProvider.<SendAsync>d__18.MoveNext()
   at Microsoft.Graph.BaseRequest.<SendRequestAsync>d__40.MoveNext()
   at Microsoft.Graph.BaseRequest.<SendAsync>d__34`1.MoveNext()
   at Microsoft.Graph.DriveItemRequest.<GetAsync>d__5.MoveNext()
   at Program.<Main>$(String[] args) in C:\Projects2022Plus\SharePointGraphAPI\SharePointGraphClientSecret\Program.cs:line 29

*/

// This creates a new folder
var taskDriveLetter = graphClient.Sites[site.Id].Lists["Letter"].Drive.Root.Children.Request().AddAsync(newFolder);
var driveLetter = taskDriveLetter.GetAwaiter().GetResult();

// This gets the folder within a speciifc Document library
var taskSearch = graphClient.Sites[site.Id].Lists["Letter"].Drive.Root.Children[newFolderName].Request().GetAsync();
var searchResult = taskSearch.GetAwaiter().GetResult();


var taskNewFolder = graphClient.Sites[site.Id].Drives[searchResult.Id].Request().GetAsync();
var lastCreatedFolder = taskNewFolder.GetAwaiter().GetResult();

Console.WriteLine($"Site Id {site.Id}");
Console.WriteLine("Please press any key to exit");
Console.ReadKey();