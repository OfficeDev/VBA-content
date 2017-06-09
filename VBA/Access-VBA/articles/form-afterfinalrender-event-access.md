---
title: Form.AfterFinalRender Event (Access)
keywords: vbaac10.chm13681
f1_keywords:
- vbaac10.chm13681
ms.prod: access
api_name:
- Access.Form.AfterFinalRender
ms.assetid: 89f9cbb5-f002-4783-dc70-17878763e486
ms.date: 06/08/2017
---


# Form.AfterFinalRender Event (Access)

Occurs after all elements in the specified PivotChart view have been rendered.


## Syntax

 _expression_. **AfterFinalRender**( ** _drawObject_** )

 _expression_ A variable that represents a **Form** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _drawObject_|Required|**Object**|A  **ChChartDraw** object. Use the methods and properties of this object to draw objects on the chart.|

### Return Value

nothing


## Example

The following example demonstrates the syntax for a subroutine that traps the  **AfterFinalRender** event.


```vb
Private Sub Form_AfterFinalRender(ByVal drawObject As Object) 
 MsgBox "The PivotChart View has fully rendered." 
End Sub
```


## See also


#### Concepts


[Form Object](form-object-access.md)

