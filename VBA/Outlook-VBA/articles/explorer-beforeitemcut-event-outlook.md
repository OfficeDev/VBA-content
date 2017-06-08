---
title: Explorer.BeforeItemCut Event (Outlook)
keywords: vbaol11.chm462
f1_keywords:
- vbaol11.chm462
ms.prod: outlook
api_name:
- Outlook.Explorer.BeforeItemCut
ms.assetid: 82861e5e-e990-aed9-4134-db9cbe63d47c
ms.date: 06/08/2017
---


# Explorer.BeforeItemCut Event (Outlook)

Occurs when an Outlook item is cut from a folder.


## Syntax

 _expression_ . **BeforeItemCut**( **_Cancel_** )

 _expression_ A variable that represents an **Explorer** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True** , the operation is not completed and the item is not deleted.|

## Remarks

This event can be cancelled after it has started. If the event is canceled, then the item will not be removed.


## Example

The following Microsoft Visual Basic for Applications (VBA) example prompts the user with a warning message before the item is cut from the folder. If the user clicks  **Yes**, the item is cut from the folder. If the user clicks  **No**, the item will not be removed from the folder. The sample code must be placed in a class module such as  `ThisOutlookSession`, and the  `Initialize_handler` routine must be called before the event procedure can be called by Microsoft Outlook.


```vb
Public WithEvents myOlExp As Outlook.Explorer 
 
Sub Initalize_Handler() 
Set myOlExp = Application.ActiveExplorer 
End Sub 
 
Private Sub myOlExp_BeforeItemCut(Cancel As Boolean) 
'Prompts the user before cutting an item 
 
 Dim lngAns As Long 
 'Display question to user 
 lngAns = MsgBox("Are you sure you want to cut the item?", vbYesNo) 
 'Set cancel argument based on user's answer 
 If lngAns = vbYes Then 
 Cancel = False 
 ElseIf lngAns = vbNo Then 
 Cancel = True 
 End If 
 
End Sub
```


## See also


#### Concepts


[Explorer Object](explorer-object-outlook.md)

