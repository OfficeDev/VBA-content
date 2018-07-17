---
title: Options.ShowScreenTipsOnObjects Property (Publisher)
keywords: vbapb10.chm1048608
f1_keywords:
- vbapb10.chm1048608
ms.prod: publisher
api_name:
- Publisher.Options.ShowScreenTipsOnObjects
ms.assetid: b5503200-31fd-72ac-de28-ace55a7123b3
ms.date: 06/08/2017
---


# Options.ShowScreenTipsOnObjects Property (Publisher)

 **True** for Microsoft Publisher to display ScreenTips when the mouse pointer hovers over a text box, shape or other object. Read/write **Boolean**.


## Syntax

 _expression_. **ShowScreenTipsOnObjects**

 _expression_A variable that represents a  **Options** object.


### Return Value

Boolean


## Example

This example disables displaying ScreenTips on objects.


```vb
Sub DisableScreenTips() 
 Options.ShowScreenTipsOnObjects = False 
End Sub
```


