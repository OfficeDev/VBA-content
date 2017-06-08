---
title: Window.Move Method (Publisher)
keywords: vbapb10.chm262163
f1_keywords:
- vbapb10.chm262163
ms.prod: publisher
api_name:
- Publisher.Window.Move
ms.assetid: a33b213b-6549-abf7-0217-041b469b798a
ms.date: 06/08/2017
---


# Window.Move Method (Publisher)

Moves the active document window.


## Syntax

 _expression_. **Move**( **_Left_**,  **_Top_**)

 _expression_A variable that represents a  **Window** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Left|Required| **Long**|The horizontal screen position of the specified window.|
|Top|Required| **Long**|The vertical screen position of the specified window.|

## Remarks

If the application window is either maximized or minimized, this method will return an error.


## Example

This example checks the state of the application window, and if it is neither maximized nor minimized, moves the window to the upper left corner of the screen.


```vb
Sub MoveWindow() 
 With ActiveWindow 
 If .WindowState = pbWindowStateNormal Then 
 .Move Left:=50, Top:=50 
 End If 
 End With 
End Sub
```


