---
title: AllForms.Parent Property (Access)
keywords: vbaac10.chm12679
f1_keywords:
- vbaac10.chm12679
ms.prod: access
api_name:
- Access.AllForms.Parent
ms.assetid: fa16ed80-9eb2-7bd8-fdc6-a8c9a8eb7ea0
ms.date: 06/08/2017
---


# AllForms.Parent Property (Access)

Returns the parent object for the specified object. Read-only.


## Syntax

 _expression_. **Parent**

 _expression_ A variable that represents an **AllForms** object.


## Remarks

You can use the  **Parent** property to determine which form or report is currently the parent when you have a subform or subreport that has been inserted in multiple forms or reports.

For example, you might insert an OrderDetails subform into both a form and a report. The following example uses the  **Parent** property to refer to the OrderID field, which is present on the main form and report. You can enter this expression in a bound control on the subform.




```
=Parent!OrderID
```


## See also


#### Concepts


[AllForms Collection](allforms-object-access.md)

