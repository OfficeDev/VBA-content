---
title: CalloutFormat.Border Property (Publisher)
keywords: vbapb10.chm2490628
f1_keywords:
- vbapb10.chm2490628
ms.prod: publisher
api_name:
- Publisher.CalloutFormat.Border
ms.assetid: 64a72ec7-4cc8-f0c7-9858-45e97bac0411
ms.date: 06/08/2017
---


# CalloutFormat.Border Property (Publisher)

Returns or sets an  **MsoTriState**constant indicating whether the text in the specified callout is surrounded by a border. Read/write.


## Syntax

 _expression_. **Border**

 _expression_A variable that represents a  **CalloutFormat** object.


### Return Value

MsoTriState


## Remarks

The  **Border** property value can be one of the ** [MsoTriState](http://msdn.microsoft.com/library/2036cfc9-be7d-e05c-bec7-af05e3c3c515%28Office.15%29.aspx)** constants declared in the Microsoft Office type library.


## Example

This example adds an oval to the active publication and a callout that points to the oval. The callout text will have a border, but not a vertical accent bar that separates the text from the callout line.


```vb
With ActiveDocument.Pages(1).Shapes 
 ' Add an oval. 
 .AddShape Type:=msoShapeOval, _ 
 Left:=180, Top:=200, Width:=280, Height:=130 
 
 ' Add a callout. 
 With .AddCallout(Type:=msoCalloutTwo, _ 
 Left:=420, Top:=170, Width:=170, Height:=40) 
 
 ' Add text to the callout. 
 .TextFrame.TextRange.Text = "This is an oval" 
 
 ' Add an accent bar to the callout. 
 With .Callout 
 .Accent = msoFalse 
 .Border = msoTrue 
 End With 
 End With 
End With 

```


