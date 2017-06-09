---
title: Application.Resize Method (Word)
keywords: vbawd10.chm158335337
f1_keywords:
- vbawd10.chm158335337
ms.prod: word
api_name:
- Word.Application.Resize
ms.assetid: 6614a0d8-eb2a-01fc-eeb6-4f8abc510bf8
ms.date: 06/08/2017
---


# Application.Resize Method (Word)

Sizes the Word application window or the specified task window.


## Syntax

 _expression_ . **Resize**( **_Width_** , **_Height_** )

 _expression_ Required. A variable that represents an **[Application](application-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Width_|Required| **Long**|The width of the window, in points.|
| _Height_|Required| **Long**|The height of the window, in points.|

## Remarks

If the window is maximized or minimized, an error occurs. Use the  **Width** or **Height** property to set the window width and height independently.


## Example

This example resizes the Microsoft Excel application window to 6 inches wide by 4 inches high.


```vb
If Tasks.Exists("Microsoft Excel") = True Then 
 With Tasks("Microsoft Excel") 
 .WindowState = wdWindowStateNormal 
 .Resize Width:=InchesToPoints(6), Height:=InchesToPoints(4) 
 End With 
End If
```

This example resizes the Word application window to 7 inches wide by 6 inches high.




```vb
With Application 
 .WindowState = wdWindowStateNormal 
 .Resize Width:=InchesToPoints(7), Height:=InchesToPoints(6) 
End With
```


## See also


#### Concepts


[Application Object](application-object-word.md)

