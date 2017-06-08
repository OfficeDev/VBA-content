---
title: Master.DrawRectangle Method (Visio)
keywords: vis_sdr.chm10716220
f1_keywords:
- vis_sdr.chm10716220
ms.prod: visio
api_name:
- Visio.Master.DrawRectangle
ms.assetid: e41ec411-ccd7-0fe6-f560-cf3934d18b59
ms.date: 06/08/2017
---


# Master.DrawRectangle Method (Visio)

Adds a rectangle to the  **Shapes** collection of a page, master, or group.


## Syntax

 _expression_ . **DrawRectangle**( **_x1_** , **_y1_** , **_x2_** , **_y2_** )

 _expression_ A variable that represents a **Master** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _x1_|Required| **Double**|The  _x_-coordinate of one corner of the rectangle's width-height box.|
| _y1_|Required| **Double**|The  _y_-coordinate of one corner of the rectangle's width-height box.|
| _x2_|Required| **Double**|The  _x_-coordinate of the other corner of the rectangle's width-height box.|
| _y2_|Required| **Double**|The  _y_-coordinate of the other corner of the rectangle's width-height box.|

### Return Value

Shape


## Remarks

Using the  **DrawRectangle** method is equivalent to using the **Rectangle** tool in the application. The arguments are in internal drawing units with respect to the coordinate space of the page, master, or group where the rectangle is being placed.


## Example

The following example shows how to draw a rectangle on the active page.


```vb
 
Public Sub DrawRectangle_Example() 
 
 Dim vsoShape As Visio.Shape 
 
 Set vsoShape = ActivePage.DrawRectangle(1, 4, 4, 1) 
 
End Sub
```


