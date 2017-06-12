---
title: Restrict Data to a Subset of Records
keywords: vbaac10.chm5187973
f1_keywords:
- vbaac10.chm5187973
ms.prod: access
ms.assetid: f3310e1f-9987-785a-9ae2-a2eb375a35c2
ms.date: 06/08/2017
---


# Restrict Data to a Subset of Records

When working with records you will often need to restrict your data to a specific set of records. Some procedures take a  _criteria_ argument that enables you to specify what data should be returned. For example, you specify the _criteria_ argument to restrict which records are returned when you use domain aggregate functions . You may also specify criteria when you use the Find method of a **Recordset** object, set the **Filter** or **ServerFilter** property of a form, or construct a[SQL Statement](build-sql-statements-that-include-variables-and-controls.md). Although each of these operations involves a different syntax, you construct the criteria expression in a similar manner for each.

For example, you can use the  **DSum** function, a domain aggregate function, to find the sum total of all freight costs in the Orders table. You could create a calculated control by entering the following expression in the **ControlSource** property:



```
=DSum("[Freight]", "Orders")
```

If you specify the optional  _criteria_ argument, the **DSum** function will be performed on a subset of _domain_. For example, you could find the sum total of all freight costs in the Orders table for only those orders being shipped to California:



```
=DSum("[Freight]", "Orders", "[ShipRegion] = 'CA'")
```

When you supply a  _criteria_ argument, Access first evaluates any expressions included in the argument to form a string expression. Then the string expression is passed to the domain function. The string expression is equivalent to an SQL WHERE clause, without the word WHERE.
You can specify numeric, textual, or date/time criteria. No matter what type of criteria you specify, the  _criteria_ argument is always converted to a string before being passed to the domain aggregate function. Therefore, you must make certain that after any expressions have been evaluated, all parts of the argument are concatenated into a single string, the whole of which is enclosed in double quotation marks (").
Use caution when constructing criteria to ensure that the string will be properly concatenated.
The following list of topics outlines the different ways in which you can specify criteria:
[Numeric Criteria Expressions](numeric-criteria-expressions.md)
[Textual Criteria Expressions](textual-criteria-expressions.md)
[Date and Time Criteria Expressions](date-and-time-criteria-expressions.md)
[Change Numeric Criteria from a Control on a Form](numeric-criteria-from-a-control-on-a-form.md)
[Change Textual Criteria from a Control on a Form](textual-criteria-from-a-control-on-a-form.md)
[Change Date and Time Criteria from a Control on a Form](date-and-time-criteria-from-a-control-on-a-form.md)
[Multiple Fields in Criteria Expressions](multiple-fields-in-criteria-expressions.md)

