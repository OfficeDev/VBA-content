---
title: Explorer.FolderSwitch Event (Outlook)
keywords: vbaol11.chm450
f1_keywords:
- vbaol11.chm450
ms.prod: outlook
api_name:
- Outlook.Explorer.FolderSwitch
ms.assetid: 5dfa1fa3-c381-8e19-0528-d70a6fd63187
ms.date: 06/08/2017
---


# Explorer.FolderSwitch Event (Outlook)

Occurs when the explorer goes to a new folder, either as a result of user action or through program code. 


## Syntax

 _expression_ . **FolderSwitch**

 _expression_ A variable that represents an **Explorer** object.


## Remarks

This event is not available in Microsoft Visual Basic Scripting Edition (VBScript).


## Example

The following Microsoft Visual Basic for Applications (VBA) example displays the  **Inbox** folder in "Messages" view whenever the user switches to the **Inbox** folder. The sample code must be placed in a class module, and the `Initialize_handler` routine must be called before the event procedure can be called by Microsoft Outlook.


```vb
Public WithEvents myOlExp As Outlook.Explorer 
 
 
 
Public Sub Initialize_handler() 
 
 Set myOlExp = Application.ActiveExplorer 
 
End Sub 
 
 
 
Private Sub myOlExp_FolderSwitch() 
 
 Select Case myOlExp.CurrentFolder.Name 
 
 Case "Inbox" 
 
 myOlExp.CurrentView = "Messages" 
 
 Case Else 
 
 End Select 
 
End Sub
```


## See also


#### Concepts


[Explorer Object](explorer-object-outlook.md)

