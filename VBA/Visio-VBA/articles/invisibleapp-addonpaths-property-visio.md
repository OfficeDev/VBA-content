---
title: InvisibleApp.AddonPaths Property (Visio)
keywords: vis_sdr.chm17513055
f1_keywords:
- vis_sdr.chm17513055
ms.prod: visio
api_name:
- Visio.InvisibleApp.AddonPaths
ms.assetid: a6709892-abc9-7043-ca51-f1b74fdb424c
ms.date: 06/08/2017
---


# InvisibleApp.AddonPaths Property (Visio)

Gets or sets the paths where Microsoft Visio looks for third-party or user add-ons. Read/write.


## Syntax

 _expression_ . **AddonPaths**

 _expression_ A variable that represents an **InvisibleApp** object.


### Return Value

String


## Remarks

The  **AddonPaths** property is set to an empty string ("") by default.

To indicate more than one folder in the path where you want Visio to look for third-party or user add-ons, use semicolons to separate individual items in the path string.

The string passed to and received from the  **AddonPaths** property is the same string shown in the **File Locations** dialog box. (Click the **File** tab, click **Options**, click  **Advanced**, and then, under  **General**, click  **File Locations**.) This string is stored in the  **HKEY_CURRENT_USER\Software\Microsoft\Office\14.0\Visio\Application\AddonsPath** subkey.

When Visio looks for third-party and user add-ons, it looks in all paths named in the  **AddonPaths** property, as well as at the paths of any add-ons installed at setup, and all the subfolders of those paths. If you pass the **AddonPaths** property to the **EnumDirectories** method, it returns a complete list of fully qualified paths in the folders passed in.

Starting with Microsoft Office Visio 2003, instead of specifying file paths to your Visio add-ons, you can publish your add-ons by using a Microsoft Windows Installer package. By doing so, you can take advantage of Microsoft Office features such as language switching, installation on demand, and repair. For more information about using a Windows Installer package to publish your add-ons, search for "Microsoft Windows Installer" on MSDN.


 **Note**  Modifying the registry in any manner, whether in the Registry Editor or programmatically, always carries some degree of risk. Incorrect modification can cause serious problems that may require you to reinstall your operating system. It is a good practice to always back up a computer's registry first before modifying it. If you are running Microsoft Windows NT or Microsoft Windows 2000, you should also update your Emergency Repair Disk (ERD).


