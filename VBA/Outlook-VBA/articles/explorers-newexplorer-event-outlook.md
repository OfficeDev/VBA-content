---
title: Explorers.NewExplorer Event (Outlook)
keywords: vbaol11.chm306
f1_keywords:
- vbaol11.chm306
ms.prod: outlook
api_name:
- Outlook.Explorers.NewExplorer
ms.assetid: 701409f3-ead3-2707-b3de-baa053e8d5c2
ms.date: 06/08/2017
---


# Explorers.NewExplorer Event (Outlook)

Occurs whenever a new explorer window is opened, either as a result of user action or through program code. 


## Syntax

 _expression_ . **NewExplorer**( **_Explorer_** )

 _expression_ A variable that represents an **Explorers** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Explorer_|Required| **[Explorer](explorer-object-outlook.md)**|The explorer that was opened.|

## Remarks

This event is not available in Microsoft Visual Basic Scripting Edition (VBScript).

The event occurs after the new  **[Explorer](explorer-object-outlook.md)** object is created but before the explorer window appears.


## Example

This Microsoft Visual Basic for Applications (VBA) example minimizes the currently active explorer window when a new explorer is about to appear. The sample code must be placed in a class module, and the  `Initialize_handler` routine must be called before the event procedure can be called by Microsoft Outlook.


```vb
Public WithEvents myOlExplorers As Outlook.Explorers 
 
 
 
Public Sub Initialize_handler() 
 
 Set myOlExplorers = Application.Explorers 
 
End Sub 
 
 
 
Private Sub myOlExplorers_NewExplorer(ByVal Explorer As Outlook.Explorer) 
 
 If TypeName(Application.ActiveExplorer) <> "Nothing" Then 
 
 Application.ActiveExplorer.WindowState = olMinimized 
 
 End If 
 
End Sub
```


## See also


#### Concepts


[Explorers Object](explorers-object-outlook.md)

