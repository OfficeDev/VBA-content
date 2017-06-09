---
title: Explorer.Deactivate Event (Outlook)
keywords: vbaol11.chm454
f1_keywords:
- vbaol11.chm454
ms.prod: outlook
api_name:
- Outlook.Explorer.Deactivate
ms.assetid: 7bf07653-3e12-670b-c293-1d51cf30e564
ms.date: 06/08/2017
---


# Explorer.Deactivate Event (Outlook)

Occurs when an explorer stops being the active window, either as a result of user action or through program code.


## Syntax

 _expression_ . **Deactivate**

 _expression_ A variable that represents an **Explorer** object.


## Remarks

This event is not available in Microsoft Visual Basic Scripting Edition (VBScript).


## Example

This Visual Basic for Applications (VBA) example uses the  **[WindowState](explorer-windowstate-property-outlook.md)** property to minimize the topmost explorer window when it is not active. The sample code must be placed in a class module, and the `Initialize_handler` routine must be called before the event procedure can be called by Outlook.


```vb
Public WithEvents myOlExp As Outlook.Explorer 
 
 
 
Public Sub Initialize_handler() 
 
 Set myOlExp = Application.ActiveExplorer 
 
End Sub 
 
 
 
Private Sub myOlExp_Deactivate() 
 
 myOlExp.WindowState = olMinimized 
 
End Sub
```


## See also


#### Concepts


[Explorer Object](explorer-object-outlook.md)

