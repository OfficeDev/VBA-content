---
title: Form.OnConnect Event (Access)
keywords: vbaac10.chm13667
f1_keywords:
- vbaac10.chm13667
ms.prod: access
api_name:
- Access.Form.OnConnect
ms.assetid: 39966052-0e06-bde9-142f-ee74d16a9973
ms.date: 06/08/2017
---


# Form.OnConnect Event (Access)

Occurs when the specified PivotTable view connects to a data source.


## Syntax

 _expression_. **OnConnect**

 _expression_ A variable that represents a **Form** object.


### Return Value

nothing


## Example

The following example demonstrates the syntax for a subroutine that traps the  **OnConnect** event.


```vb
Private Sub Form_OnConnect() 
 MsgBox "The PivotTable View has " _ 
 &; "connected to its data source!" 
End Sub
```


## See also


#### Concepts


[Form Object](form-object-access.md)

