---
title: Explorer.BeforeMinimize Event (Outlook)
keywords: vbaol11.chm458
f1_keywords:
- vbaol11.chm458
ms.prod: outlook
api_name:
- Outlook.Explorer.BeforeMinimize
ms.assetid: 999b2bc3-99de-6dc8-81a2-dd25c8bc71c6
ms.date: 06/08/2017
---


# Explorer.BeforeMinimize Event (Outlook)

Occurs when the active explorer is minimized by the user.


## Syntax

 _expression_ . **BeforeMinimize**( **_Cancel_** )

 _expression_ A variable that represents an **Explorer** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True** , the operation is not completed and the explorer is not minimized.|

## Remarks

This event can be cancelled after it has started.


## Example

The following Microsoft Visual Basic for Applications (VBA) example prompts the user with a message before the window is minimized. If the user clicks  **Yes**, the explorer is minimized. The sample code must be placed in a class module such as  `ThisOutlookSession`, and the  `Initialize_Handler()` subroutine should be called before the event procedure can be called by Microsoft Outlook.


```vb
Public WithEvents myOlExp As Outlook.Explorer 
 
 
 
Sub Initalize_Handler() 
 
 Set myOlExp = Application.ActiveExplorer 
 
End Sub 
 
 
 
Private Sub myOlExp_BeforeMinimize(Cancel As Boolean) 
 
'Prompts the user before minimizing the Explorer 
 
 
 
 Dim lngAns As Long 
 
 
 
 lngAns = MsgBox("Are you sure you want to minimize the current window?", vbYesNo) 
 
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

