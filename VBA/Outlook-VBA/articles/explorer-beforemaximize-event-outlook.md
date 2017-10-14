---
title: Explorer.BeforeMaximize Event (Outlook)
keywords: vbaol11.chm457
f1_keywords:
- vbaol11.chm457
ms.prod: outlook
api_name:
- Outlook.Explorer.BeforeMaximize
ms.assetid: 4d55aa87-44c6-4660-c2bf-579d3b9dc376
ms.date: 06/08/2017
---


# Explorer.BeforeMaximize Event (Outlook)

Occurs when an explorer is maximized by the user.


## Syntax

 _expression_ . **BeforeMaximize**( **_Cancel_** )

 _expression_ A variable that represents an **Explorer** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True** , the operation is not completed and the explorer is not maximized.|

## Remarks

This event can be cancelled after it has started.


## Example

The following Microsoft Visual Basic for Applications (VBA) example prompts the user with a warning message before maximizing the current window. If the user clicks  **Yes**, the explorer will maximize. The sample code must be placed in a class module such as  `ThisOutlookSession`, and the  `Initialize_Handler()` subroutine should be called before the event procedure can be called by Microsoft Outlook.


```vb
Public WithEvents myOlExp As Outlook.Explorer 
 
 
 
Sub Initalize_Handler() 
 
 Set myOlExp = Application.ActiveExplorer 
 
End Sub 
 
 
 
Private Sub myOlExp_BeforeMaximize(Cancel As Boolean) 
 
'Prompts the user before maximizing the explorer 
 
 
 
 Dim lngAns As Long 
 
 
 
 lngAns = MsgBox("Are you sure you want to maximize the current window?", vbYesNo) 
 
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

