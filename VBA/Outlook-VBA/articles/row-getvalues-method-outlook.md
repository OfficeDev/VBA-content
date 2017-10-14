---
title: Row.GetValues Method (Outlook)
keywords: vbaol11.chm2244
f1_keywords:
- vbaol11.chm2244
ms.prod: outlook
api_name:
- Outlook.Row.GetValues
ms.assetid: 1f92e0ab-9ba8-9cc6-51e8-05cc145a93bf
ms.date: 06/08/2017
---


# Row.GetValues Method (Outlook)

Obtains a one-dimensional array containing the values for all columns at the  **[Row](row-object-outlook.md)** in the parent **[Table](table-object-outlook.md)** .


## Syntax

 _expression_ . **GetValues**

 _expression_ A variable that represents a **Row** object.


### Return Value

A  **Variant** that represents an array of values for all the columns at that **Row** in the **Table** .


## Remarks

 **GetValues** is a helper method that facilitates fetching all the column values in the **Row** in a single call.

Since the array is zero-based, the length of the array is the number of columns in the  **Row** minus one.

Values returned in the array are of the same type as the values in the parent  **Table** . This means that binary properties in the **Table** are returned as arrays of bytes. For date-time properties, if a **[Column](column-object-outlook.md)** is a default column or if it has been added using an explicit built-in property name, then its value in the **Table** and in the array are expressed in local time. If the **Column** has been added to the **Table** using a namespace reference, then its value in the **Table** and in the array are expressed in Coordinated Universal Time (UTC). For more information on referencing properties by namespace, see[Referencing Properties by Namespace](http://msdn.microsoft.com/library/c1c7bfa9-64d7-81d2-84e7-f0a4c57780b3%28Office.15%29.aspx). 


## See also


#### Concepts


[Row Object](row-object-outlook.md)

