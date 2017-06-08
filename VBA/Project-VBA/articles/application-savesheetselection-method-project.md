---
title: Application.SaveSheetSelection Method (Project)
keywords: vbapj.chm2095
f1_keywords:
- vbapj.chm2095
ms.prod: project-server
api_name:
- Project.Application.SaveSheetSelection
ms.assetid: ed792b68-7af2-2b42-9f92-eb77e3b1780e
ms.date: 06/08/2017
---


# Application.SaveSheetSelection Method (Project)

Saves row and column information for a selected sheet view.


## Syntax

 _expression_. **SaveSheetSelection**

 _expression_ A variable that represents an **Application** object.


### Return Value

 **Boolean**


## Example

The following example demonstrates how  **SaveSheetSelection** and **RestoreSheetSelection** work.


```vb
Sub SelectionDemo() 
 
 '1) In your sheet view, make column/row/cell selections, then run this 
 '2) macro which toggles the Project Guide display state, and 
 ' clears the ActiveSelection (saved via Application.SaveSheetSelection). 
 '3) The macro then restores the ActiveSelection via Application.RestoreSheetSelection 
 
 'Save the ActiveSelection in the active sheet view 
 Application.SaveSheetSelection 
 
 'Toggle the Project Guide display state 
 Dim boolPGON As Boolean 
 boolPGON = Application.DisplayProjectGuide 
 
 If boolPGON = True Then 
 Application.DisplayProjectGuide = False 
 Else 
 Application.DisplayProjectGuide = True 
 End If 
 
 MsgBox "The Project Guide display state has been toggled. " _ 
 &; "Notice that your active selection was cleared in the " _ 
 &; "process." &; Chr(10) &; Chr(10) _ 
 &; "Now the call to RestoreSheetSelection restores the ActiveSelection...", _ 
 vbOKOnly 
 
 Application.RestoreSheetSelection 
 
End Sub
```


