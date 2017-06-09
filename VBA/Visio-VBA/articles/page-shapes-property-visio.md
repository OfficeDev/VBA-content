---
title: Page.Shapes Property (Visio)
keywords: vis_sdr.chm10914330
f1_keywords:
- vis_sdr.chm10914330
ms.prod: visio
api_name:
- Visio.Page.Shapes
ms.assetid: b6a5c174-c1d6-049b-8aec-8337c47341d7
ms.date: 06/08/2017
---


# Page.Shapes Property (Visio)

Returns the  **Shapes** collection for a page, master, or group. Read-only.


## Syntax

 _expression_ . **Shapes**

 _expression_ A variable that represents a **Page** object.


### Return Value

Shapes


## Remarks

If your Visual Studio solution includes the  **Microsoft.Office.Interop.Visio** reference, this property maps to the following types:


-  **Microsoft.Office.Interop.Visio.IVPage.Shapes**
    

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


