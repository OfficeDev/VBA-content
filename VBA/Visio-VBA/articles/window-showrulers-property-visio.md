---
title: Window.ShowRulers Property (Visio)
keywords: vis_sdr.chm11614375
f1_keywords:
- vis_sdr.chm11614375
ms.prod: visio
api_name:
- Visio.Window.ShowRulers
ms.assetid: 857dc23b-3687-2b52-db6e-358d32a422fa
ms.date: 06/08/2017
---


# Window.ShowRulers Property (Visio)

Determines whether rulers are shown in the drawing window. Read/write.


## Syntax

 _expression_ . **ShowRulers**

 _expression_ A variable that represents a **Window** object.


### Return Value

Integer


## Remarks

Setting the  **ShowRulers** property is the same as selecting **Ruler** in the **Show/Hide** group on the **View** tab.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **ShowRulers** property to switch display of the rulers on and off.


```vb
 
Public Sub ShowRulers_Example() 
 
 'Check whether the active window is a drawing window. 
 If ActiveWindow.Type = visDrawing Then 
 
 'Switch the rulers on or off. 
 ActiveWindow.ShowRulers = Not ActiveWindow.ShowRulers 
 
 Else 
 
 'Tell the user why you are not switching the rulers. 
 MsgBox "Active window is not a drawing window.", _ 
 vbOKOnly, "Show/Hide Rulers" 
 
 End If 
 
End Sub
```


