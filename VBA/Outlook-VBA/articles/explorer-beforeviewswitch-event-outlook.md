---
title: Explorer.BeforeViewSwitch Event (Outlook)
keywords: vbaol11.chm453
f1_keywords:
- vbaol11.chm453
ms.prod: outlook
api_name:
- Outlook.Explorer.BeforeViewSwitch
ms.assetid: 5b7ac070-ba4d-6fa8-94e5-20370efe7343
ms.date: 06/08/2017
---


# Explorer.BeforeViewSwitch Event (Outlook)

Occurs before the explorer changes to a new view, either as a result of user action or through program code. 


## Syntax

 _expression_ . **BeforeViewSwitch**( **_NewView_** , **_Cancel_** )

 _expression_ A variable that represents an **Explorer** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _NewView_|Required| **Variant**|The name of the view the explorer is switching to.|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True** , the switch is cancelled and the current view is not changed.|

## Remarks

This event is not available in Microsoft Visual Basic Scripting Edition (VBScript).


## Example

This Microsoft Visual Basic for Applications (VBA) example confirms that the user wants to switch views and cancels the switch if the user answers No. The sample code must be placed in a class module, and the  `Initialize_handler` routine must be called before the event procedure can be called by Microsoft Outlook.


```vb
Public WithEvents myOlExp As Outlook.Explorer 
 
 
 
Public Sub Initialize_handler() 
 
 Set myOlExp = Application.ActiveExplorer 
 
End Sub 
 
 
 
Private Sub myOlExp_BeforeViewSwitch(ByVal NewView As Variant, Cancel As Boolean) 
 
 Dim Prompt As String 
 
 
 
 Prompt = "Are you sure you want to switch to the " &; NewView &; " view?" 
 
 If MsgBox(Prompt, vbYesNo + vbQuestion) = vbNo Then Cancel = True 
 
End Sub
```


## See also


#### Concepts


[Explorer Object](explorer-object-outlook.md)

