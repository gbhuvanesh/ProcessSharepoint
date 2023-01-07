//// See https://aka.ms/new-console-template for more information
using SharePointGraphClientSecret;

Console.WriteLine("Hello, World!");


var graphClient = (new GraphCServicelientHelper()).GetGraphClient();

//var me = await graphClient.Me.Request().GetAsync();

var task = graphClient.Sites["root"].SiteWithPath("/sites/CRMDEV").Request().GetAsync();
var site = task.GetAwaiter().GetResult();
Console.WriteLine($"Site Id {site.Id}");
Console.WriteLine("Please press any key to exit");
Console.ReadKey();