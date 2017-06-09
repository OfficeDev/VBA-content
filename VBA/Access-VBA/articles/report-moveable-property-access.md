---
title: Report.Moveable Property (Access)
keywords: vbaac10.chm13811
f1_keywords:
- vbaac10.chm13811
ms.prod: access
api_name:
- Access.Report.Moveable
ms.assetid: 77e682a5-7a0f-f55e-a469-2770bb2de844
ms.date: 06/08/2017
---


# Report.Moveable Property (Access)

Returns or sets a  **Boolean** indicating whether the specified report can be moved by the user; **True** if it can be moved. Read/write.


## Syntax

 _expression_. **Moveable**

 _expression_ A variable that represents a **Report** object.


## Remarks

You can use the  **Move** method to programmatically move a form or report regardless of the value of the **Moveable** property.


## Example

The following example determines whether or not the first form in the current project can be moved.


```vb
If Forms(0).Moveable Then 
 MsgBox "You may move the form." 
Else 
 MsgBox "The form cannot be moved." 
End If 

```


## See also


#### Concepts


[Report Object](report-object-access.md)

