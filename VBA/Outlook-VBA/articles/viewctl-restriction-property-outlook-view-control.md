---
title: ViewCtl.Restriction Property (Outlook View Control)
ms.prod: outlook
ms.assetid: 5e92a492-653d-27f1-8d3e-799987d911be
ms.date: 06/08/2017
---


# ViewCtl.Restriction Property (Outlook View Control)

Returns or sets a  **String** value that represents a filter to the items that are displayed in the control. As a result, the control displays only those items that match the filter. Read/write.


## Syntax

 _expression_. **Restriction**

 _expression_A variable that represents a  **ViewCtl** object.


## Remarks

The filter is a string expression that contains one or more filter clauses that are joined by the logical operators  **Or**,  **Not**, or  **And**.

A filter clause is a simple expression that evaluates to  **True** or **False**; for example,  `[CompanyName] = "Microsoft"`

Note that property names can be used in the expression and are identified and delimited by square brackets. Except for these bracketed property names, only literals are allowed within the expression; variables and constants are not evaluated as part of the expression.

Comparison operators allowed within the filter expression include >, <, >=, <=, = and <>. Comparisons are not case sensitive and do not include subject prefixes that are added when a message is replied to or forwarded. 

Note that "=" is not interpreted as "equals" in  **String** comparisons, but as "contains" instead, so that `[Subject] = 'Outlook'` matches all items that have "Outlook" or "outlook" anywhere in the Subject field. To create a true equality filter, you must use <= and >= together, as in the following example.




```
OvCtl1.Restriction "[Subject] <= 'outlook'
```

and 




```
[Subject] >= 'outlook'
```

In this example, the control displays only those items whose Subject field contains only "outlook" or "Outlook."

The setting of the  **Restriction** property does not persist if the view or current folder changes.

The  **Restriction** property only works correctly if you use Table or Card views. This is a limitation of the Microsoft Outlook View Control.


