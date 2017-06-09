---
title: Explorer.SelectionChange Event (Outlook)
keywords: vbaol11.chm455
f1_keywords:
- vbaol11.chm455
ms.prod: outlook
api_name:
- Outlook.Explorer.SelectionChange
ms.assetid: ef0d976f-b9f6-2080-7657-e48d1c64ccb1
ms.date: 06/08/2017
---


# Explorer.SelectionChange Event (Outlook)

Occurs when the user selects a different or additional Microsoft Outlook item programmatically or by interacting with the user interface.


## Syntax

 _expression_ . **SelectionChange**

 _expression_ A variable that represents an **[Explorer](explorer-object-outlook.md)** object.


## Remarks

This event also occurs when the user (either programmatically or via the user interface) clicks or switches to a different folder that contains items, because Outlook automatically selects the first item in that folder. However, this event does not occur if the folder is a file-system folder or if any folder with a current Web view is displayed. This event is not available in Microsoft Visual Basic Scripting Edition (VBScript).


## Example

The following Microsoft Visual Basic for Applications (VBA) example displays the number of items that are selected in the active explorer window whenever the selection changes. The sample code must be placed in a class module, and the  `Initialize_handler` routine must be called before the event procedure can be called by Microsoft Outlook.


```vb
Public WithEvents myOlExp As Outlook.Explorer 
 
 
 
Public Sub Initialize_handler() 
 
 Set myOlExp = Application.ActiveExplorer 
 
End Sub 
 
 
 
Private Sub myOlExp_SelectionChange() 
 
 MsgBox myOlExp.Selection.Count &; " items selected." 
 
End Sub
```


## See also


#### Concepts


[Explorer Object](explorer-object-outlook.md)

