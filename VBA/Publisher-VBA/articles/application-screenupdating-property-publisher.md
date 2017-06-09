---
title: Application.ScreenUpdating Property (Publisher)
keywords: vbapb10.chm131107
f1_keywords:
- vbapb10.chm131107
ms.prod: publisher
api_name:
- Publisher.Application.ScreenUpdating
ms.assetid: d265b4fb-1452-91a5-32fe-0cad54c8f29c
ms.date: 06/08/2017
---


# Application.ScreenUpdating Property (Publisher)

Returns or sets a  **Boolean** indicating whether Microsoft Publisher refreshes the screen display during run time; **True** to refresh the screen display. Read/write.


## Syntax

 _expression_. **ScreenUpdating**

 _expression_A variable that represents a  **Application** object.


### Return Value

Boolean


## Remarks

Turning screen updating off during run time can speed execution of Microsoft Visual Basic code. However, we recommend that you provide some indication of status so that the user is aware that the program is functioning correctly.


## Example

The following example turns off screen updating at the beginning of a subroutine and turns it back on at the end of the subroutine.


```vb
Sub TurnOffScreenUpdating() 
 ScreenUpdating = False 
 
 ' Execute code here. 
 
 ScreenUpdating = True 
End Sub
```


## See also


#### Concepts


 [Application Object](application-object-publisher.md)

