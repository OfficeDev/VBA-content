---
title: Shapes.AddCallout Method (Word)
keywords: vbawd10.chm161415178
f1_keywords:
- vbawd10.chm161415178
ms.prod: word
api_name:
- Word.Shapes.AddCallout
ms.assetid: 5745edcc-5010-8df8-5311-9179461e01fe
ms.date: 06/08/2017
---


# Shapes.AddCallout Method (Word)

Adds a borderless line callout to a drawing canvas. .


## Syntax

 _expression_ . **AddCallout**( **_Type_** , **_Left_** , **_Top_** , **_Width_** , **_Height_** )

 _expression_ Required. A variable that represents a **[Shapes](shapes-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Type_|Required| **MsoCalloutType**|The type of callout.|
| _Left_|Required| **Single**|The position, in points, of the left edge of the callout's bounding box.|
| _Top_|Required| **Single**|The position, in points, of the top edge of the callout's bounding box.|
| _Width_|Required| **Single**|The width, in points, of the callout's bounding box.|
| _Height_|Required| **Single**|The height, in points, of the callout's bounding box.|

### Return Value

 **[Shape](shape-object-word.md)**


## Remarks

You can insert a greater variety of callouts, such as balloons and clouds, using the  **AddShape** method.


## Example

This example adds a callout to a newly created drawing canvas.


```vb
Sub NewCanvasCallout() 
 Dim shpCanvas As Shape 
 
 'Add drawing canvas to the active document 
 Set shpCanvas = ActiveDocument.Shapes.AddCanvas _ 
 (Left:=150, Top:=150, Width:=200, Height:=300) 
 
 'Add callout to the drawing canvas 
 shpCanvas.CanvasItems.AddCallout _ 
 Type:=msoCalloutTwo, Left:=100, _ 
 Top:=40, Width:=150, Height:=75 
End Sub
```


## See also


#### Concepts


[Shapes Collection Object](shapes-object-word.md)

