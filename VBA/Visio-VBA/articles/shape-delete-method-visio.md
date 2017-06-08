---
title: Shape.Delete Method (Visio)
keywords: vis_sdr.chm11216165
f1_keywords:
- vis_sdr.chm11216165
ms.prod: visio
api_name:
- Visio.Shape.Delete
ms.assetid: 0960d9e1-b091-ea8c-0724-e10a68d8821a
ms.date: 06/08/2017
---


# Shape.Delete Method (Visio)

Deletes an object or selection.


## Syntax

 _expression_ . **Delete**

 _expression_ A variable that represents a **Shape** object.


### Return Value

Nothing


## Remarks

If your Visual Studio solution includes the  **Microsoft.Office.Interop.Visio** reference, this method maps to the following types:


-  **Microsoft.Office.Interop.Visio.IVShape.Delete()**
    

## Example

This example shows how to delete  **Shape** objects.


```vb
 
Public Sub Delete_Example()  
 
    Dim vsoShape As Visio.Shape  
 
    Set vsoShape = ActivePage.DrawLine(1, 1, 5, 5)  
 
    vsoShape.Delete  
 
End Sub
```


