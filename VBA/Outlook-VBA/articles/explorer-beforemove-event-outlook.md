---
title: Explorer.BeforeMove Event (Outlook)
keywords: vbaol11.chm459
f1_keywords:
- vbaol11.chm459
ms.prod: outlook
api_name:
- Outlook.Explorer.BeforeMove
ms.assetid: bce617d3-3bf8-2a59-ab0a-4ef1e7759c75
ms.date: 06/08/2017
---


# Explorer.BeforeMove Event (Outlook)

Occurs when the  **[Explorer](explorer-object-outlook.md)** is moved by the user.


## Syntax

 _expression_ . **BeforeMove**( **_Cancel_** )

 _expression_ A variable that represents an **Explorer** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True** , the operation is not completed and the explorer or inspector is not moved.|

## Remarks

This event can be cancelled after it has started.


## Example

The following Microsoft Visual Basic for Applications (VBA) example prompts the user with a message before the explorer is moved by the user. If the user clicks  **Yes**, the explorer can be moved by the user. The sample code must be placed in a class module such as  `ThisOutlookSession`, and the  `Initialize_Handler()` subroutine should be called before the event procedure can be called by Microsoft Outlook.


```vb
Public WithEvents myOlExp As Outlook.Explorer 
 
 
 
Sub Initalize_Handler() 
 
 Set myOlExp = Application.ActiveExplorer 
 
End Sub 
 
 
 
Private Sub myOlExp_BeforeMove(Cancel As Boolean) 
 
'Prompts the user before moving the window 
 
 
 
 Dim lngAns As Long 
 
 
 
 lngAns = MsgBox("Are you sure you want to move the current window? Use your keyboard to make your selection.", vbYesNo) 
 
 If lngAns = vbYes Then 
 
 Cancel = False 
 
 Else 
 
 Cancel = True 
 
 End If 
 
 
 
End Sub
```


## See also


#### Concepts


[Explorer Object](explorer-object-outlook.md)

