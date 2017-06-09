---
title: Folder.Name Property (Outlook)
keywords: vbaol11.chm1991
f1_keywords:
- vbaol11.chm1991
ms.prod: outlook
api_name:
- Outlook.Folder.Name
ms.assetid: ec03a345-8c06-f234-e1e9-ecdc54495ed2
ms.date: 06/08/2017
---


# Folder.Name Property (Outlook)

Returns or sets a  **String** value that represents the display name for the object. Read/write.


## Syntax

 _expression_ . **Name**

 _expression_ A variable that represents a **Folder** object.


## Example

This Visual Basic for Applications (VBA) example uses the  **Name** property to obtain the name of the folder displayed in the active explorer.


```vb
Sub DisplayCurrentFolderName() 
 
 Dim myExplorer As Outlook.Explorer 
 
 Dim myFolder As Outlook.Folder 
 
 
 
 Set myExplorer = Application.ActiveExplorer 
 
 Set myFolder = myExplorer.CurrentFolder 
 
 MsgBox myFolder.Name 
 
End Sub
```


## See also


#### Concepts


[Folder Object](folder-object-outlook.md)

