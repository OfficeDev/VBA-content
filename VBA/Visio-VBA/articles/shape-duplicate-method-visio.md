---
title: Shape.Duplicate Method (Visio)
keywords: vis_sdr.chm11216245
f1_keywords:
- vis_sdr.chm11216245
ms.prod: visio
api_name:
- Visio.Shape.Duplicate
ms.assetid: a45fd247-e4ad-8149-3656-af9588f076ef
ms.date: 06/08/2017
---


# Shape.Duplicate Method (Visio)

Duplicates an object.


## Syntax

 _expression_ . **Duplicate**

 _expression_ A variable that represents a **Shape** object.


### Return Value

Shape


## Remarks

The  **Duplicate** method duplicates the specified object or selection and adds a copy to the same page as the original. Using the **Duplicate** method is equivalent to clicking **Duplicate** on the **Paste** menu on the **Home** tab.

When used with a  **Shape** object, the **Duplicate** method duplicates the shape.

When used with a  **Selection** object, the **Duplicate** method duplicates the selection.


## Example

The following example shows how to duplicate  **Shape** objects. The code also works for **Selection** objects.

Before running this macro, make sure a drawing page is active in the Microsoft Visio window.




```vb
 
Public Sub Duplicate_Example() 
 
 Dim vsoOriginalShape As Visio.Shape 
 Dim vsoDuplicateShape As Visio.Shape 
 
 Set vsoOriginalShape = ActivePage.DrawLine(1, 1, 5, 5) 
 
 Set vsoDuplicateShape = vsoOriginalShape.Duplicate 
 vsoDuplicateShape.Cells("BeginY") = "2" 
 
End Sub
```


