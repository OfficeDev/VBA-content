---
title: Application.StartupPaths Property (Visio)
keywords: vis_sdr.chm10014415
f1_keywords:
- vis_sdr.chm10014415
ms.prod: visio
api_name:
- Visio.Application.StartupPaths
ms.assetid: 966a91d9-9ada-d0e1-9886-271ea47faaf9
ms.date: 06/08/2017
---


# Application.StartupPaths Property (Visio)

Gets or sets the paths where Microsoft Visio looks for third-party and user add-ons to run when the application is started. Read/write.


## Syntax

 _expression_ . **StartupPaths**

 _expression_ A variable that represents an **Application** object.


### Return Value

String


## Remarks

The  **StartupPaths** property is set to an empty string ("") by default.

The string passed to and received from the  **StartupPaths** property is the same string shown in the **File Locations** dialog box. (Click the **File** tab, click **Options**, click  **Advanced**, and then, under  **General**, click  **File Locations**.) This string is stored in the  **HKEY_CURRENT_USER\Software\Microsoft\Office\14.0\Visio\Application\StartupPath** subkey.

When the application looks for third-party and user startup add-on files, it looks in all paths named in the  **StartupPaths** property, as well as at the paths of any startup add-ons installed at setup, and all the subfolders of those paths. If you pass the **StartupPaths** property to the **EnumDirectories** method, it returns a complete list of fully qualified paths in the folders passed in.

Setting the  **StartupPaths** property replaces existing values for **StartupPaths** in the **File Locations** dialog box. To retain existing values, get the existing string and then append the new file path to that string, as shown in the following code:




```vb
Application.StartupPaths = Application.StartupPaths &; ";" &; "newpath ".
```


 **Caution**  Modifying the Microsoft Windows registry in any manner, whether in the Registry Editor or programmatically, always carries some degree of risk. Incorrect modification can cause serious problems that may require you to reinstall your operating system. It is a good practice to always back up a computer's registry first before modifying it. If you are running Microsoft Windows NT or Microsoft Windows 2000, you should also update your Emergency Repair Disk (ERD).


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **StartupPaths** property to add a path to the **Start-up** paths list.


```vb
Public Sub StartupPaths_Example() 
  
    Dim strMessage As String 
    Dim strNewPath As String 
    Dim strStartupPath As String 
    Dim strTitle As String  
 
    'Get the path we want to add.  
    strStartupPath = Application.StartupPaths  
    strTitle = "StartupPaths"  
    strMessage = "The current content of the Visio Start-up paths box is:"  
    strMessage = strMessage &; vbCrLf &; strStartupPath  
    MsgBox strMessage, vbInformation + vbOKOnly, strTitle  
    strMessage = "Type in an additional path for Visio to look for add-ons. "  
         
    strNewPath = InputBox$(strMessage, strTitle)  
 
    'Make sure the folder exists and that it's not 
    'already in the Start-up paths box.  
    strMessage = ""  
 
    If strNewPath = ""  Then 
        strMessage = "You did not enter a path." 
    ElseIf InStr(strStartupPath, strNewPath)  Then 
        strMessage = "The path you specified is already in the Start-up paths box." 
    ElseIf Len(Dir$(strNewPath, vbDirectory)) = 0 And _  
                Len(Dir$(Application.Path &; strNewPath, _  
                vbDirectory)) = 0 Then 
        strMessage = "The folder you typed does not exist (or is empty)." 
    Else 
        Application.StartupPaths = strStartupPath &; ";" &; strNewPath 
        strMessage = "We just added " &; strNewPath &; _  
                " to the startup paths." 
    End If 
       
    If strMessage <> ""  Then 
        MsgBox strMessage, vbExclamation + vbOKOnly, strTitle  
    End If 
  
End Sub
```


