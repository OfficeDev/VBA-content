---
title: Report.Painting Property (Access)
keywords: vbaac10.chm13736
f1_keywords:
- vbaac10.chm13736
ms.prod: access
api_name:
- Access.Report.Painting
ms.assetid: 82c5a5e6-9d87-7293-e0f5-8ee950f3b85f
ms.date: 06/08/2017
---


# Report.Painting Property (Access)

You can use the  **Painting** property to specify whether a report is repainted. Read/write **Boolean**.


## Syntax

 _expression_. **Painting**

 _expression_ A variable that represents a **Report** object.


## Remarks

The  **Painting** property is similar to the Echo action. However, the **Painting** property prevents repainting of a single report, whereas the Echo action prevents repainting of all open windows in an application.

Setting the  **Painting** property for a report to **False** also prevents all controls (except subreport controls) on a report from being repainted. To prevent a subreport control from being repainted, you must set the **Painting** property for the subreport to **False**. (Note that you set the **Painting** property for the subreport, not the subreport control.)

The  **Painting** property is automatically set to **True** whenever the report gets or loses the focus. You can set this property to **False** while you are working on a report if you don't want to see changes to the report or to its controls. For example, if a form has a set of controls that are automatically resized when the form is resized and you don't want the user to see each individual control move, you can turn **Painting** off, and then move all of the controls, then turn **Painting** back on.


## Example

The following example uses the  **Painting** property to enable or disable form painting depending on whether the `SetPainting` variable is set to **True** or **False**. If form painting is turned off, Microsoft Access displays the hourglass icon while painting is turned off.


```vb
Public Sub EnablePaint(ByRef frmName As Form, _ 
 ByVal SetPainting As Integer) 
 
 frmName.Painting = SetPainting 
 
 ' Form painting is turned off. 
 If SetPainting = False Then 
 DoCmd.Hourglass True 
 Else 
 DoCmd.Hourglass False 
 End If 
 
End Sub
```


## See also


#### Concepts


[Report Object](report-object-access.md)

