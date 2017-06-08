---
title: ShapeRange.Callout Property (Excel)
keywords: vbaxl10.chm640099
f1_keywords:
- vbaxl10.chm640099
ms.prod: excel
api_name:
- Excel.ShapeRange.Callout
ms.assetid: 15078411-7968-27ba-aa73-2c5d69220b08
ms.date: 06/08/2017
---


# ShapeRange.Callout Property (Excel)

Returns a  **[CalloutFormat](calloutformat-object-excel.md)** object that contains callout formatting properties for the specified shape. Applies to a **[ShapeRange](shaperange-object-excel.md)** object that represent line callouts. Read-only.


## Syntax

 _expression_ . **Callout**

 _expression_ A variable that represents a **ShapeRange** object.


## Example

This example adds to  `myDocument` an oval and a callout that points to the oval. The callout text won't have a border, but it will have a vertical accent bar that separates the text from the callout line.


```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes 
 .AddShape msoShapeOval, 180, 200, 280, 130 
 With .AddCallout(msoCalloutTwo, 420, 170, 170, 40) 
 .TextFrame.Characters.Text = "My oval" 
 With .Callout 
 .Accent = True 
 .Border = False 
 End With 
 End With 
End With
```


## See also


#### Concepts


[ShapeRange Object](shaperange-object-excel.md)

