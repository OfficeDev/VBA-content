---
title: Window.ShowGrid Property (Visio)
keywords: vis_sdr.chm11614350
f1_keywords:
- vis_sdr.chm11614350
ms.prod: visio
api_name:
- Visio.Window.ShowGrid
ms.assetid: 288e1b14-5ad5-da14-8f5b-747212093247
ms.date: 06/08/2017
---


# Window.ShowGrid Property (Visio)

Determines whether a grid is shown in a window. Read/write.


## Syntax

 _expression_ . **ShowGrid**

 _expression_ A variable that represents a **Window** object.


### Return Value

Integer


## Remarks

Setting the  **ShowGrid** property is equivalent to selecting **Grid** in the **Show/Hide** group on the **View** tab.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **ShowGrid** property to hide the grid. To restore the grid after running this macro, set the **ShowGrid** property to **True** .


```vb
 
Public Sub ShowGrid_Example() 
 
 'Check whether active window is a drawing window. 
 If ActiveWindow.Type = visDrawing Then 
 
 'Hide the grid. 
 ActiveWindow.ShowGrid = False 
 
 Else 
 
 'Tell the user why you are not hiding the grid. 
 MsgBox "Active window is not a drawing window.", vbOKOnly, "Show Grid" 
 
 End If 
 
End Sub
```


