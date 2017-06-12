---
title: SyncObject.Progress Event (Outlook)
keywords: vbaol11.chm112
f1_keywords:
- vbaol11.chm112
ms.prod: outlook
api_name:
- Outlook.SyncObject.Progress
ms.assetid: 605c0243-45c1-94d9-8356-b31bb1d0d3e1
ms.date: 06/08/2017
---


# SyncObject.Progress Event (Outlook)

Occurs periodically while Microsoft Outlook is synchronizing a user?s folders using the specified  **Send\Receive** group.


## Syntax

 _expression_ . **Progress**( **_State_** , **_Description_** , **_Value_** , **_Max_** )

 _expression_ A variable that represents a **SyncObject** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _State_|Required| **[OlSyncState](olsyncstate-enumeration-outlook.md)**|A value that identifies the current state of the synchronization process.|
| _Description_|Required| **String**|A textual description of the current state of the synchronization process.|
| _Value_|Required| **Long**|Specifies the current value of the synchronization process (such as the number of items synchronized).|
| _Max_|Required| **Long**|The maximum that  _Value_ can reach. The ratio of _Value_ to _Max_ represents the percent complete of the synchronization process.|

## Remarks

This event is not available in Microsoft Visual Basic Scripting Edition (VBScript).


## Example

This Microsoft Visual Basic for Applications (VBA) example shows the progress of synchronization. The sample code must be placed in a class module, and the  `Initialize_handler` routine must be called before the event procedure can be called by Outlook.


```vb
Public WithEvents mySync As Outlook.SyncObject 
 
Sub Initialize_handler() 
 Set mySync = Application.Session.SyncObjects.Item(1) 
 mySync.Start 
End Sub 
 
Private Sub mySync_Progress(ByVal State As Outlook.OlSyncState, ByVal Description As String, ByVal Value As Long, ByVal Max As Long) 
 If Not Description = "" Then 
 MsgBox Description 
 End If 
End Sub
```


## See also


#### Concepts


[SyncObject Object](syncobject-object-outlook.md)

