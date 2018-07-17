---
title: ShapeRange.ZOrder Method (Excel)
keywords: vbaxl10.chm640095
f1_keywords:
- vbaxl10.chm640095
ms.prod: excel
api_name:
- Excel.ShapeRange.ZOrder
ms.assetid: 3a2e8556-ddbf-312d-85a3-6cd5d2865499
ms.date: 06/08/2017
---


# ShapeRange.ZOrder Method (Excel)

Moves the specified shape in front of or behind other shapes in the collection (that is, changes the shape's position in the z-order).


## Syntax

 _expression_ . **ZOrder**( **_ZOrderCmd_** )

 _expression_ A variable that represents a **ShapeRange** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ZOrderCmd_|Required| **[MsoZOrderCmd](http://msdn.microsoft.com/library/4615d1c7-9d7b-70a4-1821-785c3af11f4f%28Office.15%29.aspx)**|Specifies where to move the specified shape relative to the other shapes.|

## Remarks



| **MsoZOrderCmd** can be one of these **MsoZOrderCmd** constants.|
| **msoBringForward**|
| **msoBringInFrontOfText** . Used only in Microsoft Word.|
| **msoBringToFront**|
| **msoSendBackward**|
| **msoSendBehindText** . Used only in Microsoft Word.|
| **msoSendToBack**|
Use the  **[ZOrderPosition](shaperange-zorderposition-property-excel.md)** property to determine a shape's current position in the z-order.


## Example

This example adds an oval to  `myDocument` and then places the oval second from the back in the z-order if there is at least one other shape on the document.


```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes.AddShape(msoShapeOval, 100, 100, 100, 300) 
    While .ZOrderPosition > 2 
        .ZOrder msoSendBackward 
    Wend 
End With
```


## See also


#### Concepts


[ShapeRange Object](shaperange-object-excel.md)

