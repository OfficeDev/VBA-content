---
title: NameSpace.RemoveStore Method (Outlook)
keywords: vbaol11.chm772
f1_keywords:
- vbaol11.chm772
ms.prod: outlook
api_name:
- Outlook.NameSpace.RemoveStore
ms.assetid: 4353387a-0e44-1d4a-b0e6-96e2c2594a6d
ms.date: 06/08/2017
---


# NameSpace.RemoveStore Method (Outlook)

Removes a Personal Folders file (.pst) from the current MAPI profile or session.


## Syntax

 _expression_ . **RemoveStore**( **_Folder_** )

 _expression_ A variable that represents a **NameSpace** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Folder_|Required| **[Folder](folder-object-outlook.md)**|The Personal Folders file (.pst) to be deleted from the list.|

## Remarks

This method removes a store only from the Microsoft Outlook user interface. You cannot remove a store from the main mailbox on the server or from a user's hard disk using the Outlook object model.


## Example

The following Microsoft Visual Basic for Applications (VBA) examples removes a folder called Personal Folders from the list of folders.


```vb
Sub RemovePST() 
 
 Dim objName As Outlook.NameSpace 
 
 Dim objFolder As Outlook.Folder 
 
 
 
 Set objName = Application.GetNamespace("MAPI") 
 
 Set objFolder = objName.Folders.Item("Personal Folders") 
 
 'Prompt the user for confirmation 
 
 Dim strPrompt As String 
 
 strPrompt = "Are you sure you want to remove the Personal Folders file?" 
 
 If MsgBox(strPrompt, vbYesNo + vbQuestion) = vbYes Then 
 
 objName.RemoveStore objFolder 
 
 End If 
 
End Sub
```


## See also


#### Concepts


[NameSpace Object](namespace-object-outlook.md)

