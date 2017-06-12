---
title: LineFormat.InsetPen Property (Publisher)
keywords: vbapb10.chm3408148
f1_keywords:
- vbapb10.chm3408148
ms.prod: publisher
api_name:
- Publisher.LineFormat.InsetPen
ms.assetid: 955b152d-517f-b5fa-6e23-765ddeb41d46
ms.date: 06/08/2017
---


# LineFormat.InsetPen Property (Publisher)

Returns or sets an  **MsoTriState** constant indicating whether a specified shape's lines are drawn inside its boundaries. Read/write.


## Syntax

 _expression_. **InsetPen**

 _expression_A variable that represents an  **LineFormat** object.


### Return Value

MsoTriState


## Remarks

An error occurs if you attempt to set this property to  **msoTrue** for any Microsoft Office AutoShape that does not support inset pen drawing.

The value of the  **InsetPen** property for tables is always **msoTrue**; attempting to set the property to any other value results in an error.

The  **InsetPen** property value can be one of the **MsoTriState** constants declared in the Microsoft Office type library and shown in the following table.



|**Constant**|**Description**|
|:-----|:-----|
| **msoFalse**|Lines are drawn directly on the specified shape's boundaries.|
| **msoTriStateMixed**|Return value indicating a combination of  **msoTrue** and **msoFalse** for the specified shape range.|
| **msoTriStateToggle**|Set value that switches between  **msoTrue** and **msoFalse**.|
| **msoTrue**|Lines are drawn inside the specified shape's boundaries.|

## Example

The following example adds two rectangles to page one of the active publication, the first with its lines drawn inside its boundaries, and the second with its lines drawn on its boundaries.


```vb
Dim shpNew As Shape 
 
With ActiveDocument.Pages(1).Shapes 
 Set shpNew = .AddShape(Type:=msoShapeRectangle, _ 
 Left:=200, Top:=150, Width:=150, Height:=100) 
 With shpNew.Line 
 .Weight = 24 
 .InsetPen = msoTrue 
 End With 
 
 Set shpNew = .AddShape(Type:=msoShapeRectangle, _ 
 Left:=200, Top:=300, Width:=150, Height:=100) 
 With shpNew.Line 
 .Weight = 24 
 .InsetPen = msoFalse 
 End With 
End With
```


