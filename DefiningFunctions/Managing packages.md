# Managing packages

QueryStorm comes with NuGet support, which it uses for two purposes:

- Installing NuGet pacakages into projects
- Sharing QueryStorm extension packages

## NuGet packages

The package manager is used for managing NuGet packages in projects (viewing, installing, uninstalling, updating). 

![PackageManager](../Images/PackageManager.png)

Installing a package into your project does the following:
- Adds an entry in the `Dependencies` section in the `module.cfg` file,
- Downloads the package
- Unpacks the contents of the package into the project
  
All dll files contained inside the package are unpacked into the `lib` folder and are automatically referenced.

Since the package contents get stored inside the project, installing NuGet packages into a workbook project will increase the size of the workbook file. Most NuGet packages are fairly small, however, so the increase in size is unlikely to be an issue. The advantage to this is that workbook applications are self-contained and will run on any machine that has the QueryStorm runtime.

## QueryStorm extensions

QueryStorm also uses its NuGet infrastructure for publishing and downloading extension packages. 

Extension packages are created and published from the QueryStom IDE (by package creators), but are installed and used from the QueryStorm runtime (by end users).

Extension packages are normal NuGet packages, but their content is designed to be used by the QueryStorm Runtime. They usually consist of one or more dlls, and an `application.manifest` file that the QueryStorm IDE generates. The manifest file tells the QueryStorm Runtime what the package contains (i.e. list of included functions) and where its entry point is.

A typical scenario for QueryStorm extensions is as follows: 

1. A developer or a consultant uses QueryStorm IDE to write a set of Excel functions (using C#, VB.NET or SQL) 
2. Once they've prepared the functions, they publish the package
3. The end users download the package via the "Extensions" button in the QueryStorm Runtime ribbon
4. The end users then uses the new functions in Excel workbooks

## Managing NuGet sources

When publishing a package, you must tell QueryStorm where to publish the package:

![Publish to feed](../images/PublishToFeed.png)

In order to see your packages, your users must add your package source to their list:

![Edit package sources](../images/EditPackageSources.png)

1. Area for managing sources (feeds)
2. Button for editing the feed
3. Feed url or path
4. Feed content type (Packages, Extensions or Both)

> Both creators and consumers use the PackageManagerDialog -> ManageSources screen to edit their package feeds.

### Types of feeds

If you are distributing packages inside your network, a shared network folder will be a good place to store packages. 

In cases where you want to distribute packages to users outside of your local network (e.g. to your clients), you can publish the package to an online NuGet server, like Azure Artifacts ([instructions](../todo)).

When adding a feed, you can specify if the source contains regular NuGet packages, QueryStorm extension packages or both. If you specify that the source contains only regular NuGet packages, it will not be included when searching for QueryStorm extensions, and vice-versa. 

Out of the box, two package sources are included:

- Nuget.org
- QueryStorm Official Feed 

Nuget.org is the main repository of .NET packages and is used in QueryStorm for installing packages into projects. 

The QueryStorm Official Feed contains only QueryStorm extensions. It's purpose is to be a place all users can download general-purpose extensions from. For now, users cannot publish to this source. 

For distributing packages to your own clients, you should use a network share or create your own Azure Artifacts repository ([instructions](../todo)). Your clients should add your feed into their list in order to be able to install your packages.

## Publishing extension packages

To publish a package, follow these steps:
1. Build the project
2. Right-click on the project and select "Publish"
3. Choose the feed to publish to
4. Enter information about the package

![Publish context menu](../Images/PublishContextMenu.png)

![Publish dialog](../Images/PublishDialog.png)

When publishing a new version of an existing package, make sure to **increment the version number**, otherwise the server will report a collision with the version that is already on the server. If the repository is a network share, though, there will be no version checks. 

## Publishing to Azure artifacts

Azure Artifacts is a cloud package management solution that allows you to create and share NuGet packages via feeds that can be both public and private to an organization with teams of any size.

Setting up a feed takes just a few minutes, and is free of charge. Currently, the free plan allows for up to 2GB of storage, which is enough for thousands of packages. Should you need more space, scaling it is quite easy as well. 

To create an Azure artifacts feed, follow the steps below:

1. Go to http://dev.azure.com/ and create an account
2. Create a new project and make it public
3. Select the Artifacts tab
4. Click "Create Feed" and give the feed a name
5. Click "Connect to Feed"
6. Click "Visual Studio" and copy the source link
7. Go to QueryStorm in Excel and add a new package source with the url from the previous step

In order to be able to publish to this feed, you'll also need to set up a personal access token. To do so, follow these steps:

1. Click on the user settings in the top right corner of the page (in the azure webpage)
2. Select "Personal Access Token"
3. Set token expiration date
4. Grant the token full access
5. Copy the token
6. In QueryStorm, enter the token as the password of the new feed and your email address (that's associated with your azure account) as the username

That's it. Now you can publish to your new feed! Share the feed URL with your users, and you're good to go.

For a video of the entire process, please click below:

[![Setting up Azure Artifacts feed](../images/video.jpg)](https://youtu.be/jc5l4OV0PZM "Setting up Azure Artifacts feed")