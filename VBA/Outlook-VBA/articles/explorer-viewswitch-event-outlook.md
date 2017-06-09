---
title: Explorer.ViewSwitch Event (Outlook)
keywords: vbaol11.chm452
f1_keywords:
- vbaol11.chm452
ms.prod: outlook
api_name:
- Outlook.Explorer.ViewSwitch
ms.assetid: ab981f42-d429-ccd7-a25c-142e52683020
ms.date: 06/08/2017
---


# Explorer.ViewSwitch Event (Outlook)

Occurs when the view in the explorer changes, either as a result of user action or through program code. 


## Syntax

 _expression_ . **ViewSwitch**

 _expression_ A variable that represents an **Explorer** object.


## Remarks

This event is not available in Microsoft Visual Basic Scripting Edition (VBScript).


## Example

This Visual Basic for Applications (VBA) example hides the preview pane if it is visible when the user switches to Messages with AutoPreview view. The sample code must be placed in a class module, and the  `Initialize_handler` routine must be called before the event procedure can be called by Microsoft Outlook.


```vb
Dim WithEvents myOlExpl As Outlook.Explorer 
 
 
 
Sub Initialize_handler() 
 
 Set myOlExpl = Application.ActiveExplorer 
 
End Sub 
 
 
 
Private Sub myOlExpl_ViewSwitch() 
 
 If myOlExpl.CurrentView = "Messages with AutoPreview" And myOlExpl.IsPaneVisible(olPreview) = True Then 
 
 myOlExpl.ShowPane olPreview, False 
 
 End If 
 
End Sub
```


## See also


#### Concepts


[Explorer Object](explorer-object-outlook.md)

