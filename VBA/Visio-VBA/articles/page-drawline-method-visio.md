---
title: Page.DrawLine Method (Visio)
keywords: vis_sdr.chm10916200
f1_keywords:
- vis_sdr.chm10916200
ms.prod: visio
api_name:
- Visio.Page.DrawLine
ms.assetid: a03308a6-7ad0-ecaa-d15d-a243402c8bd3
ms.date: 06/08/2017
---


# Page.DrawLine Method (Visio)

Adds a line to the  **Shapes** collection of a page.


## Syntax

 _expression_ . **DrawLine**( **_xBegin_** , **_yBegin_** , **_xEnd_** , **_yEnd_** )

 _expression_ A variable that represents a **Page** object.


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


