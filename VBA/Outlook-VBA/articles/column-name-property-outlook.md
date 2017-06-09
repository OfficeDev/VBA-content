---
title: Column.Name Property (Outlook)
keywords: vbaol11.chm2749
f1_keywords:
- vbaol11.chm2749
ms.prod: outlook
api_name:
- Outlook.Column.Name
ms.assetid: e69a8a53-d348-2147-28cf-d41ea80bba61
ms.date: 06/08/2017
---


# Column.Name Property (Outlook)

Returns a  **String** value that represents the name of the **[Column](column-object-outlook.md)** . Read-only.


## Syntax

 _expression_ . **Name**

 _expression_ A variable that represents a **Column** object.


## Remarks

The  **Name** property is the default member of the **Column** object.

If the  **Column** is a default column in the **[Table](table-object-outlook.md)** , or if it has been added to the **Table** with the explicit built-in name for the property, the value of **Name** is the explicit built-in name (without any enclosing brackets) for the property. If the **Column** has been added to the **Table** with a property name referencing a namespace, the value of **Name** will be the property name referenced by namespace. For more information on referencing properties by namespace, see[Referencing Properties by Namespace](http://msdn.microsoft.com/library/c1c7bfa9-64d7-81d2-84e7-f0a4c57780b3%28Office.15%29.aspx).


## See also


#### Concepts


[Column Object](column-object-outlook.md)

