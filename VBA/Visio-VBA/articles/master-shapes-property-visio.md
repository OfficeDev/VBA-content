---
title: Master.Shapes Property (Visio)
keywords: vis_sdr.chm10714330
f1_keywords:
- vis_sdr.chm10714330
ms.prod: visio
api_name:
- Visio.Master.Shapes
ms.assetid: 56db5c02-9b55-dfe1-993b-c23e93e84577
ms.date: 06/08/2017
---


# Master.Shapes Property (Visio)

Returns the  **Shapes** collection for a page, master, or group. Read-only.


## Syntax

 _expression_ . **Shapes**

 _expression_ A variable that represents a **Master** object.


### Return Value

Shapes


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **Shapes** property to get the **Shapes** collection. It prints the names of all shapes on Page1 in the Immediate window.

To run this macro, make sure the active document has shapes on Page1.




```vb
 
Public Sub Shapes_Example() 
 
 Dim intCounter As Integer 
 Dim intShapeCount As Integer 
 Dim vsoShapes As Visio.Shapes 
 
 Set vsoShapes = ActiveDocument.Pages.Item(1).Shapes 
 
 Debug.Print "Shapes in document: "; ActiveDocument.Name 
 Debug.Print "On page: "; ActiveDocument.Pages.Item(1).Name 
 
 intShapeCount = vsoShapes.Count 
 
 If intShapeCount > 0 Then 
 For intCounter = 1 To intShapeCount 
 Debug.Print " "; vsoShapes.Item(intCounter).Name 
 Next intCounter 
 
 Else 
 Debug.Print "No Shapes On Page" 
 End If 
 
End Sub
```


