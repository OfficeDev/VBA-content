---
title: Shape.DrawOval Method (Visio)
keywords: vis_sdr.chm11216210
f1_keywords:
- vis_sdr.chm11216210
ms.prod: visio
api_name:
- Visio.Shape.DrawOval
ms.assetid: 7f561251-251e-6aa9-5291-5919ccce1a9e
ms.date: 06/08/2017
---


# Shape.DrawOval Method (Visio)

Adds an oval (ellipse) to the  **Shapes** collection of a group shape.


## Syntax

 _expression_ . **DrawOval**( **_x1_** , **_y1_** , **_x2_** , **_y2_** )

 _expression_ A variable that represents a **Shape** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _x1_|Required| **Double**|The x-coordinate of one corner of the ellipse's width-height box.|
| _y1_|Required| **Double**|The y-coordinate of one corner of the ellipse's width-height box.|
| _x2_|Required| **Double**|The x-coordinate of the other corner of the ellipse's width-height box.|
| _y2_|Required| **Double**|The y-coordinate of the other corner of the ellipse's width-height box.|

### Return Value

Shape


## Remarks

Using the  **DrawOval** method is equivalent to using the **Ellipse** tool in the application. The arguments are in internal drawing units with respect to the coordinate space of the page, master, or group where the ellipse is being placed.


## Example

The following example shows how to draw an oval (ellipse) on the active page.


```vb
 
Public Sub DrawOval_Example() 
 
 Dim vsoShape As Visio.Shape 
 
 Set vsoShape = ActivePage.DrawOval(1.5, 10.5, 7.5, 6.5) 
 
End Sub
```


