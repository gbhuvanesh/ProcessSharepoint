//// See https://aka.ms/new-console-template for more information
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

var newFolder = new DriveItem
{
    Name = "newFolder" + DateTime.Now.ToString("yyyyMMddhhmmss"),
    Folder = new Folder()
};

var taskDriveLetter = graphClient.Sites[site.Id].Lists["Letter"].Drive.Root.Children.Request().AddAsync(newFolder);
var driveLetter = taskDriveLetter.GetAwaiter().GetResult();


Console.WriteLine($"Site Id {site.Id}");
Console.WriteLine("Please press any key to exit");
Console.ReadKey();