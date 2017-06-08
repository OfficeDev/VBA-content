---
title: Application.DisplayScrollBars Property (Word)
keywords: vbawd10.chm158335058
f1_keywords:
- vbawd10.chm158335058
ms.prod: word
api_name:
- Word.Application.DisplayScrollBars
ms.assetid: 23b3957a-e4c1-b422-836a-074f84ff2f8e
ms.date: 06/08/2017
---


# Application.DisplayScrollBars Property (Word)

 **True** if Word displays a scroll bar in at least one document window. **False** if there are no scroll bars displayed in any window. Read/write **Boolean** .


## Syntax

 _expression_ . **DisplayScrollBars**

 _expression_ A variable that represents an **[Application](application-object-word.md)** object.


## Remarks

Setting the  **DisplayScrollBars** property to **True** displays horizontal and vertical scroll bars in all windows. Setting this property to **False** turns off all scroll bars in all windows.

Use the  **DisplayHorizontalScrollBar** and **DisplayVerticalScrollBar** properties to display individual scroll bars in the specified window.


## Example

This example displays horizontal and vertical scroll bars in all windows.


```vb
Application.DisplayScrollBars = True
```

This example returns True if there is a scroll bar currently displayed in any window.




```vb
Dim blnTemp As Boolean 
 
blnTemp = Application.DisplayScrollBars
```


## See also


#### Concepts


[Application Object](application-object-word.md)

