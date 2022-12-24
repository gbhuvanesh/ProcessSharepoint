// See https://aka.ms/new-console-template for more information
using SharePointGraphClientSecret;

Console.WriteLine("Hello, World!");


var graphClient = (new GraphCServicelientHelper()).GetGraphClient();

var siteId = await graphClient.Sites["root"].Request().GetAsync();

Console.WriteLine("Please press any key to exit");
Console.ReadKey();