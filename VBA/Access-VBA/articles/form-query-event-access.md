---
title: Form.Query Event (Access)
keywords: vbaac10.chm13670
f1_keywords:
- vbaac10.chm13670
ms.prod: access
api_name:
- Access.Form.Query
ms.assetid: f3070a6f-3064-b496-ff9f-4da165205f90
ms.date: 06/08/2017
---


# Form.Query Event (Access)

Occurs whenever the specified PivotTable view query becomes necessary. The query may not occur immediately; it may be delayed until the new data is displayed.


## Syntax

 _expression_. **Query**

 _expression_ A variable that represents a **Form** object.


## Example

The following example demonstrates the syntax for a subroutine that traps the  **Query** event.


```vb
Private Sub Form_Query() 
 MsgBox "A query has become necessary." 
End Sub
```


## See also


#### Concepts


[Form Object](form-object-access.md)

