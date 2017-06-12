---
title: Application.TemplatePaths Property (Visio)
keywords: vis_sdr.chm10014510
f1_keywords:
- vis_sdr.chm10014510
ms.prod: visio
api_name:
- Visio.Application.TemplatePaths
ms.assetid: 149a9ef2-e255-3dad-2177-b29c173fa66d
ms.date: 06/08/2017
---


# Application.TemplatePaths Property (Visio)

Gets or sets the paths where Microsoft Visio looks for templates. Read/write.


## Syntax

 _expression_ . **TemplatePaths**

 _expression_ A variable that represents an **Application** object.


### Return Value

 **String**


## Remarks

The  **TemplatePaths** property is set to an empty string ("") by default.

The string passed to and received from the  **TemplatePaths** property is the same string shown in the **File Locations** dialog box. (Click the **File** tab, click **Options**, click  **Advanced**, and then, under  **General**, click  **File Locations**.) This string is stored in the  **HKEY_CURRENT_USER\Software\Microsoft\Office\14.0\Visio\Application\TemplatePath** subkey.

When Visio looks for templates, it looks in all paths named in the  **TemplatePaths** property and all the subfolders of those paths. If you pass the **TemplatePaths** property to the **EnumDirectories** method, it returns a complete list of fully qualified paths in the folders passed in.

Setting the  **TemplatePaths** property replaces existing values for **Templates** in the **File Locations** dialog box. To retain existing values, get the existing string and then append the new file path to that string, as shown in the following code:




```vb
Application.TemplatePaths = Application.TemplatePaths &; ";" &; "newpath ".
```


 **Caution**  Modifying the registry in any manner, whether in the Registry Editor or programmatically, always carries some degree of risk. Incorrect modification can cause serious problems that may require you to reinstall your operating system. It is a good practice to always back up a computer's registry first before modifying it. If you are running Microsoft Windows NT or Microsoft Windows 2000, you should also update your Emergency Repair Disk (ERD). 


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **TemplatePaths** property to add a path to the **Templates** paths box.


```vb
 
Public Sub TemplatePaths_Example()  
 
    Dim strMessage As String 
    Dim strNewPath As String 
    Dim strTemplatePath As String 
    Dim strTitle As String 
     
    'Get the path we want to add.  
    strTemplatePath = Application.TemplatePaths  
    strTitle = "TemplatePaths"  
    strMessage = "The current content of the Visio Templates path box is:"  
    strMessage = strMessage &; vbCrLf &; strTemplatePath  
    MsgBox strMessage, vbInformation + vbOKOnly, strTitle  
    strMessage = "Type in an additional path for Visio to look for templates. "  
    strNewPath = InputBox$(strMessage, strTitle)  
 
    'Make sure the folder exists and that it's not 
    'already in the Templates path box.  
    strMessage = ""  
    If strNewPath = ""  Then 
        strMessage = "You did not enter a path." 
        ElseIf InStr(strTemplatePath, strNewPath)  Then 
            strMessage = "The path you specified is already in the Templates path box."  
        ElseIf Len(Dir$(strNewPath, vbDirectory)) = 0 And _  
                Len(Dir$(Application.Path &; strNewPath, _  
                vbDirectory)) = 0 Then 
            strMessage = "The folder you typed does not exist (or is blank)."  
        Else 
            Application.TemplatePaths = strTemplatePath &; ";" &; strNewPath  
            strMessage = "We just added " &; strNewPath &; _  
                " to the Templates path box."  
    End If 
   
    If strMessage <> ""  Then 
        MsgBox strMessage, vbExclamation + vbOKOnly, strTitle  
    End If  
  
End Sub
```


