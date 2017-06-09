---
title: Options.DisplayStatusBar Property (Publisher)
keywords: vbapb10.chm1048583
f1_keywords:
- vbapb10.chm1048583
ms.prod: publisher
api_name:
- Publisher.Options.DisplayStatusBar
ms.assetid: 335b2f1e-03ff-fd90-5ec2-27d5219b27e7
ms.date: 06/08/2017
---


# Options.DisplayStatusBar Property (Publisher)

 **True** for Microsoft Publisher to show the status bar at the bottom of the Publisher window. Read/write **Boolean**.


## Syntax

 _expression_. **DisplayStatusBar**

 _expression_A variable that represents a  **Options** object.


### Return Value

Boolean


## Example

This example hides the status bar from view.


```vb
Sub HideStatusBar() 
 Options.DisplayStatusBar = False 
End Sub
```


