---
title: Inspector.BeforeSize Event (Outlook)
keywords: vbaol11.chm471
f1_keywords:
- vbaol11.chm471
ms.prod: outlook
api_name:
- Outlook.Inspector.BeforeSize
ms.assetid: ee0b12af-0edc-bd06-c67c-67469df128dd
ms.date: 06/08/2017
---


# Inspector.BeforeSize Event (Outlook)

Occurs when the user sizes the current  **[Inspector](inspector-object-outlook.md)** .


## Syntax

 _expression_ . **BeforeSize**( **_Cancel_** )

 _expression_ A variable that represents an **Inspector** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True** , the operation is not completed and the **Inspector** is not sized.|

## Remarks

This event can be cancelled after it has started. If the event is cancelled, the window is not sized.


## Example

The following Microsoft Visual Basic for Applications (VBA) example prompts the user with a warning message before the inspector is sized. If the user clicks  **Yes**, the inspector can be sized. The sample code must be placed in a class module such as  `ThisOutlookSession`, and the  `Initialize_Handler()` subroutine should be called before the event procedure can be called by Microsoft Outlook.


```vb
Public WithEvents myIns As Outlook.Inspector 
 
 
 
Public Sub Initalize_Handler() 
 
 Set myIns = Application.ActiveInspector 
 
End Sub 
 
 
 
Private Sub myIns_BeforeSize(Cancel As Boolean) 
 
 'Prompts the user before resizing the window 
 
 Dim lngAns As Long 
 
 lngAns = MsgBox("Are you sure you want to resize the current window? Use your keyboard to make your selection.", vbYesNo) 
 
 If lngAns = vbYes Then 
 
 Cancel = False 
 
 Else 
 
 Cancel = True 
 
 End If 
 
End Sub
```


## See also


#### Concepts


[Inspector Object](inspector-object-outlook.md)

