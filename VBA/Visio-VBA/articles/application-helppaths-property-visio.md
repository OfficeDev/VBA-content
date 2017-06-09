---
title: Application.HelpPaths Property (Visio)
keywords: vis_sdr.chm10013635
f1_keywords:
- vis_sdr.chm10013635
ms.prod: visio
api_name:
- Visio.Application.HelpPaths
ms.assetid: eba05b64-61d8-970d-65f4-26ea41840fcf
ms.date: 06/08/2017
---


# Application.HelpPaths Property (Visio)

Gets or sets the paths where Microsoft Visio looks for Help files. Read/write.


## Syntax

 _expression_ . **HelpPaths**

 _expression_ A variable that represents an **Application** object.


### Return Value

String


## Remarks

The  **HelpPaths** property is set to an empty string ("") by default.

The string passed to and received from the  **HelpPaths** property is the same string shown in the **File Paths** dialog box. (Click the **File** tab, click **Options**, click  **Advanced**, and then, under  **General**, click ** File Locations**.) This string is stored in the  **HKEY_CURRENT_USER\Software\Microsoft\Office\14.0\Visio\Application\HelpPath** subkey.

When the application looks for Help files, it looks in all paths named in the  **HelpPaths** property and all the subfolders of those paths. If you pass the **HelpPaths** property to the **EnumDirectories** method, it returns a complete list of fully qualified paths in the folders passed in.

Setting the  **HelpPaths** property replaces existing values for **HelpPaths** in the **File Paths** dialog box. To retain existing values, get the existing string and then append the new file path to that string, as shown in the following code:




```vb
Application.HelpPaths = Application.HelpPaths &; ";" &; "newpath ".
```


 **Note**  Modifying the registry in any manner, whether through the Registry Editor or programmatically, always carries some degree of risk. Incorrect modification can cause serious problems that may require you to reinstall your operating system. It is a good practice to always back up a computer's registry first before modifying it. If you are running Microsoft Windows NT or Microsoft Windows 2000, you should also update your Emergency Repair Disk (ERD).


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to get and set the  **HelpPaths** property of the **Application** object. Before running this macro, replace _fullpath(s)_ with the path or paths to the location or locations where you want Visio to look for Help files.


```vb
 
Public Sub GetHelpPaths_Example()  
 
    Dim strCurrentPath As String 
 
    'Retrieve the current path to Visio Help files.  
    strCurrentPath = Application.HelpPaths  
    MsgBox ("The current path for Microsoft Visio Help files is " + strCurrentPath)  
 
End Sub   
 
Public Sub SetHelpPaths_Example()  
 
    Dim strNewPath As String 
 
    'Store the new path.  
    strNewPath = "fullpath(s) "  
 
    'Set the new path in the Application object.  
    Application.HelpPaths = strNewPath  
 
End Sub
```


