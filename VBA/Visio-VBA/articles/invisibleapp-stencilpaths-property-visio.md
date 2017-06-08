---
title: InvisibleApp.StencilPaths Property (Visio)
keywords: vis_sdr.chm17514440
f1_keywords:
- vis_sdr.chm17514440
ms.prod: visio
api_name:
- Visio.InvisibleApp.StencilPaths
ms.assetid: 9cbf837e-c192-78e4-b829-a33425e82f45
ms.date: 06/08/2017
---


# InvisibleApp.StencilPaths Property (Visio)

Gets or sets the paths where Microsoft Visio looks for stencils. Read/write.


## Syntax

 _expression_ . **StencilPaths**

 _expression_ A variable that represents an **InvisibleApp** object.


### Return Value

 **String**


## Remarks

The  **StencilPaths** property is set to an empty string ("") by default.

The string passed to and received from the  **StencilPaths** property is the same string shown in the **File Locations** dialog box. (Click the **File** tab, click **Options**, click  **Advanced**, and then, under  **General**, click  **File Locations**.) This string is stored in the  **HKEY_CURRENT_USER\Software\Microsoft\Office\14.0\Visio\Application\StencilPath** subkey.

When Visio looks for stencils, it looks in all paths named in the  **StencilPaths** property and all the subfolders of those paths. If you pass the **StencilPaths** property to the **EnumDirectories** method, it returns a complete list of fully qualified paths in the folders passed in.

Setting the  **StencilPaths** property replaces existing values for **Stencils** in the **File Locations** dialog box. To retain existing values, get the existing string and then append the new file path to that string, as shown in the following code:




```vb
Application.StencilPaths = Application.StencilPaths &; ";" &; "newpath ".
```


 **Caution**   Modifying the registry in any manner, whether in the Registry Editor or programmatically, always carries some degree of risk. Incorrect modification can cause serious problems that may require you to reinstall your operating system. It is a good practice to always back up a computer's registry first before modifying it. If you are running Microsoft Windows NT or Microsoft Windows 2000, you should also update your Emergency Repair Disk (ERD).


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to us the  **StencilPaths** property to add a path to the stencils.


```vb
 
Public Sub ShowStencilPaths_Example() 
  
    Dim strMessage As String 
    Dim strNewPath As String 
    Dim strStencilPath As String 
    Dim strTitle As String 
 
    'Get the path we want to add.  
    strStencilPath = Application.StencilPaths  
    strTitle = "StencilPaths"  
    strMessage = "The current content of the Visio Stencils box is:"  
    strMessage = strMessage &; vbCrLf &; strStencilPath  
    MsgBox strMessage, vbInformation + vbOKOnly, strTitle  
    strMessage = "Type in an additional path for Visio to look for stencils. "  
    strNewPath = InputBox$(strMessage, strTitle)  
 
    'Make sure the folder exists and that it's not 
    'already in the stencil paths.  
    strMessage = ""  
    If strNewPath = ""  Then 
        strMessage = "You did not enter a path." 
        ElseIf InStr(strStencilPath, strNewPath)  Then 
            strMessage = "The path you specified is already in the stencil paths."  
        ElseIf Len(Dir$(strNewPath, vbDirectory)) = 0 And _  
                Len(Dir$(Application.Path &; strNewPath, _  
                vbDirectory)) = 0 Then 
            strMessage = "The folder you typed does not exist (or is blank)."  
        Else 
            Application.StencilPaths = strStencilPath &; ";" &; strNewPath  
            strMessage = "We just added " &; strNewPath &; _  
                " to the stencil paths."  
        End If 
   
    If strMessage <> ""  Then 
        MsgBox strMessage, vbExclamation + vbOKOnly, strTitle  
    End If 
   
End Sub
```


