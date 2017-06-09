---
title: Form.DataSetChange Event (Access)
keywords: vbaac10.chm13677
f1_keywords:
- vbaac10.chm13677
ms.prod: access
api_name:
- Access.Form.DataSetChange
ms.assetid: b266f48e-ccf9-1be1-edfb-f99892b09c97
ms.date: 06/08/2017
---


# Form.DataSetChange Event (Access)

Occurs whenever the specified PivotTable view is data-bound and the data set changes — for example, when a filter operation takes place. This event also occurs when initial data is available from the data source.


## Syntax

 _expression_. **DataSetChange**

 _expression_ A variable that represents a **Form** object.


### Return Value

nothing


## Example

The following example demonstrates the syntax for a subroutine that traps the  **DataSetChange** event.


```vb
Private Sub Form_DataSetChange() 
 MsgBox "The data set for the PivotChart view has changed." 
End Sub
```


## See also


#### Concepts


[Form Object](form-object-access.md)

