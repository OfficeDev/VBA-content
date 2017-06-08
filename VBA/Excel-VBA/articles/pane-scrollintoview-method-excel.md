---
title: Pane.ScrollIntoView Method (Excel)
keywords: vbaxl10.chm360080
f1_keywords:
- vbaxl10.chm360080
ms.prod: excel
api_name:
- Excel.Pane.ScrollIntoView
ms.assetid: 650020f6-cc4a-fe19-8c7a-3c2ed9b27e16
ms.date: 06/08/2017
---


# Pane.ScrollIntoView Method (Excel)

Scrolls the document window so that the contents of a specified rectangular area are displayed in either the upper-left or lower-right corner of the document window or pane (depending on the value of the  _Start_ argument).


## Syntax

 _expression_ . **ScrollIntoView**( **_Left_** , **_Top_** , **_Width_** , **_Height_** , **_Start_** )

 _expression_ A variable that represents a **Pane** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Left_|Required| **Long**|The horizontal position of the rectangle (in points) from the left edge of the document window or pane.|
| _Top_|Required| **Long**|The vertical position of the rectangle (in points) from the top of the document window or pane.|
| _Width_|Required| **Long**|The width of the rectangle, in points.|
| _Height_|Required| **Long**|The height of the rectangle, in points.|
| _Start_|Optional| **Variant**| **True** to have the upper-left corner of the rectangle appear in the upper-left corner of the document window or pane. **False** to have the lower-right corner of the rectangle appear in the lower-right corner of the document window or pane. The default value is **True** .|

## Remarks

The  _Start_ argument is useful for orienting the screen display when the rectangle is larger than the document window or pane.


## Example

This example defines a 100-by-200-pixel rectangle in the active document window, positioned 20 pixels from the top of the window and 50 pixels from the left edge of the window.The example then scrolls the document up and to the left so that the upper-left corner of the rectangle is aligned with the upper-left corner of the window.


```vb
ActiveWindow.ScrollIntoView _ 
 Left:=50, Top:=20, _ 
 Width:=100, Height:=200
```


## See also


#### Concepts


[Pane Object](pane-object-excel.md)

