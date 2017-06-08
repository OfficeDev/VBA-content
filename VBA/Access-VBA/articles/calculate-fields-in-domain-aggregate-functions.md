---
title: Calculate Fields in Domain Aggregate Functions
keywords: vbaac10.chm5187048
f1_keywords:
- vbaac10.chm5187048
ms.prod: access
ms.assetid: 73c27d1c-0a3c-03e4-c17c-337133d7b316
ms.date: 06/08/2017
---


# Calculate Fields in Domain Aggregate Functions

You can use the string expression argument (the  _expr_ argument) in a domain aggregate function to perform a calculation on values in a field. For example, you can calculate a percentage (such as a surcharge or sales tax) by dividing a field value by a number.

The following table provides examples of calculations on fields from an Orders table and an Order Details table.


|**Calculation**|**Example**|
|:-----|:-----|
|Add a number to a field|"[Freight] + 5"|
|Subtract a number from a field|"[Freight] - 5"|
|Multiply a field by a number|"[Freight] * 2"|
|Divide a field by a number|"[Freight] / 2"|
|Add one field to another|"[UnitsInStock] + [UnitsOnOrder]"|
|Subtract one field from another|"[ReorderLevel] - [UnitsInStock]"|
You would most likely use a domain aggregate function in a macro or module, in a calculated control on a form or report, or in a criteria expression in a query.
For example, you can calculate the average discount amount for all orders in an Order Details table. Multiply the Unit Price and Discount fields to determine the discount for each order, then calculate the average. Enter the following example in a procedure in a module.



```vb
Dim dblX As Double 
dblX = DAvg("[UnitPrice] * [Discount]", "[Order Details]")
```


