---
title: Explorer.BeforeItemCopy Event (Outlook)
keywords: vbaol11.chm461
f1_keywords:
- vbaol11.chm461
ms.prod: outlook
api_name:
- Outlook.Explorer.BeforeItemCopy
ms.assetid: 05ae7be8-5528-5560-f8ce-73f0afbf4cde
ms.date: 06/08/2017
---


# Explorer.BeforeItemCopy Event (Outlook)

Occurs when an Outlook item is copied.


## Syntax

 _expression_ . **BeforeItemCopy**( **_Cancel_** )

 _expression_ A variable that represents an **Explorer** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True** , the operation is not completed and the item is not copied.|

## Remarks

This event can be cancelled after it has started.


## Example

The following Microsoft Visual Basic for Applications (VBA) example prompts the user before an item is copied. A message is displayed to the user verifying that the item should be copied. If the user clicks  **Yes**, the item is copied to the Clipboard. The sample code must be placed in a class module such as  `ThisOutlookSession`, and the  `Initialize_handler` routine must be called before the event procedure can be called by Microsoft Outlook.


```vb
Public WithEvents myOlExp As Outlook.Explorer 
 
 
 
Sub Initalize_Handler() 
 
Set myOlExp = Application.ActiveExplorer 
 
End Sub 
 
 
 
Private Sub myOlExp_BeforeItemCopy(Cancel As Boolean) 
 
'Prompts the user before copying an item 
 
 
 
 Dim lngAns As Long 'user answer 
 
 'Display question to user 
 
 lngAns = MsgBox("Are you sure you want to copy the item?", vbYesNo) 
 
 If lngAns = vbYes Then 
 
 Cancel = False 
 
 Else 
 
 'Set Cancel argument based on answer 
 
 Cancel = True 
 
 End If 
 
 
 
End Sub
```


## See also


#### Concepts


[Explorer Object](explorer-object-outlook.md)

