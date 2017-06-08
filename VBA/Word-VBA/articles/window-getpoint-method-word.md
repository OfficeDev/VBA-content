---
title: Window.GetPoint Method (Word)
keywords: vbawd10.chm157417584
f1_keywords:
- vbawd10.chm157417584
ms.prod: word
api_name:
- Word.Window.GetPoint
ms.assetid: b0f2b558-0dfc-96f8-5177-3771f6fbb69b
ms.date: 06/08/2017
---


# Window.GetPoint Method (Word)

Returns the screen coordinates of the specified range or shape.


## Syntax

 _expression_ . **GetPoint**( **_ScreenPixelsLeft_** , **_ScreenPixelsTop_** , **_ScreenPixelsWidth_** , **_ScreenPixelsHeight_** , **_obj_** )

 _expression_ Required. A variable that represents a **[Window](window-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ScreenPixelsLeft_|Required| **Long**|The variable name to which you want Microsoft Word to return the value for the left edge of the object.|
| _ScreenPixelsTop_|Required| **Long**|The variable name to which you want Word to return the value for the top edge of the object.|
| _ScreenPixelsWidth_|Required| **Long**|The variable name to which you want Word to return the value for the width of the object.|
| _ScreenPixelsHeight_|Required| **Long**|The variable name to which you want Word to return the value for the height of the object.|
| _obj_|Required| **Object**|A  **Range** or **Shape** object.|

## Remarks

If the entire range or shape isn't visible on the screen, an error occurs.


## Example

This example examines the current selection and returns its screen coordinates.


```vb
Dim pLeft As Long 
Dim pTop As Long 
Dim pWidth As Long 
Dim pHeight As Long 
 
ActiveWindow.GetPoint pLeft, pTop, pWidth, pHeight, _ 
 Selection.Range 
MsgBox "Left = " &; pLeft &; vbLf _ 
 &; "Top = " &; pTop &; vbLf _ 
 &; "Width = " &; pWidth &; vbLf _ 
 &; "Height = " &; pHeight
```


## See also


#### Concepts


[Window Object](window-object-word.md)

