---
title: Shape.DrawLine Method (Visio)
keywords: vis_sdr.chm11216200
f1_keywords:
- vis_sdr.chm11216200
ms.prod: visio
api_name:
- Visio.Shape.DrawLine
ms.assetid: 8793104a-0ded-e2ca-54e8-acf987b9c797
ms.date: 06/08/2017
---


# Shape.DrawLine Method (Visio)

Adds a line to the  **Shapes** collection of a group shape.


## Syntax

 _expression_ . **DrawLine**( **_xBegin_** , **_yBegin_** , **_xEnd_** , **_yEnd_** )

 _expression_ A variable that represents a **Shape** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _xBegin_|Required| **Double**|The x-coordinate of the line's begin point.|
| _yBegin_|Required| **Double**|The y-coordinate of the line's begin point.|
| _xEnd_|Required| **Double**|The x-coordinate of the line's endpoint.|
| _yEnd_|Required| **Double**|The y-coordinate of the line's endpoint.|

### Return Value

Shape


## Remarks

Using the  **DrawLine** method is equivalent to using the **Line** tool in Microsoft Visio. The arguments are in internal drawing units with respect to the coordinate space of the page, master, or group where the line is being placed.


## Example

The following example shows how to draw a line shape on the active page.


```vb
 
Public Sub DrawLine_Example() 
 
 Dim vsoShape As Visio.Shape 
 
 Set vsoShape = ActivePage.DrawLine(5, 4, 7.5, 1) 
 
End Sub
```


