---
title: Selection.Duplicate Method (Visio)
keywords: vis_sdr.chm11116245
f1_keywords:
- vis_sdr.chm11116245
ms.prod: visio
api_name:
- Visio.Selection.Duplicate
ms.assetid: 515b522c-8b99-ea51-822f-47f0de24d330
ms.date: 06/08/2017
---


# Selection.Duplicate Method (Visio)

Duplicates a selection.


## Syntax

 _expression_ . **Duplicate**

 _expression_ A variable that represents a **Selection** object.


### Return Value

Selection


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


