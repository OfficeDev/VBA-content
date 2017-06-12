---
title: Explorer.BeforeFolderSwitch Event (Outlook)
keywords: vbaol11.chm451
f1_keywords:
- vbaol11.chm451
ms.prod: outlook
api_name:
- Outlook.Explorer.BeforeFolderSwitch
ms.assetid: ae65c073-6b4a-ac81-c4ae-691118b19df0
ms.date: 06/08/2017
---


# Explorer.BeforeFolderSwitch Event (Outlook)

Occurs before the explorer goes to a new folder, either as a result of user action or through program code.


## Syntax

 _expression_ . **BeforeFolderSwitch**( **_NewFolder_** , **_Cancel_** )

 _expression_ A variable that represents an **Explorer** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _NewFolder_|Required| **Object**|The  **[Folder](folder-object-outlook.md)** object the explorer is switching to.|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True** , navigation is cancelled, and the current folder is not changed.|

## Remarks

This event is not available in Microsoft Visual Basic Scripting Edition (VBScript).

If the folder being switched to is in a namespace that doesn?t support automation (such as the file system),  _NewFolder_ is **Nothing** .


## Example

This sample prevents a user from switching to a folder named "Off Limits". The sample code must be placed in a class module such as  `ThisOutlookSession`, and the  `Initialize_handler` routine must be called before the event procedure can be called by Microsoft Outlook. To run this example without errors, make sure a folder by the name 'Off Limits' exists in the folder displayed in the active explorer.


```vb
Public WithEvents myOlExp As Outlook.Explorer 
 
 
 
Public Sub Initialize_handler() 
 
 Set myOlExp = Application.ActiveExplorer 
 
End Sub 
 
 
 
Private Sub myOlExp_BeforeFolderSwitch(ByVal NewFolder As Object, Cancel As Boolean) 
 
 If NewFolder.Name = "Off Limits" Then 
 
 MsgBox "You do not have permission to access this folder." 
 
 Cancel = True 
 
 End If 
 
End Sub
```


## See also


#### Concepts


[Explorer Object](explorer-object-outlook.md)

