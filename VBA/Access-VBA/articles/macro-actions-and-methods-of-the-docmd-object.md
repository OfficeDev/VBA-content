---
title: Macro Actions and Methods of the DoCmd Object
keywords: vbaac10.chm5187441
f1_keywords:
- vbaac10.chm5187441
ms.prod: access
ms.assetid: aab25fbe-8ec3-5c45-dd70-a0e8c885406a
ms.date: 06/08/2017
---


# Macro Actions and Methods of the DoCmd Object

To carry out macro actions from code in Access, use the [DoCmd](docmd-object-access.md)object and its methods. This object replaces the  **DoCmd** statement that you used in versions 1. _x_ and 2.0 of Access to carry out a macro action.

When you convert a database, Access automatically converts any  **DoCmd** statements and the actions that they carried out in your Access Basic code to methods of the **DoCmd** object by replacing the space with the . (dot) operator.

Some macro actions work differently in Access 9.0 and later than in version 1. _x_, 2.0, or 7.0; these differences are detailed below.


## Databases Created with Access 95

 _The DoMenuItem Action_

The DoMenuItem action is no longer used in Access. The RunCommand action can be used to accomplish the tasks for which you used to use the DoMenuItem action.

When you enable a database created with a prior version of Access, the DoMenuItem action will continue to work as it did before.

When you convert a database created with a prior version of Access, all DoMenuItem actions in macros are replaced with RunCommand actions the first time that the macros are saved after conversion. DoMenuItem methods used in Visual Basic procedures aren't changed.


## Databases Created with Access Version 1.

 _The TransferSpreadsheet Action_

Access can't import Excel version 2.0 spreadsheets or Lotus 1-2-3 version 1.0 spreadsheets. If your converted database contains a macro that provided this functionality by using the TransferSpreadsheet action in Access version 1. _x_ or 2.0, converting the database will change the Spreadsheet Type argument to Excel version 3.0 (if you originally specified Excel version 2.0), or causes an error if you originally specified Lotus 1-2-3 version 1.0 format.

To work around this problem, convert the spreadsheets to a later version of Excel or Lotus 1-2-3 before importing them into Access.

 _The TransferText and TransferSpreadsheet Actions_

In Access, you can't use a SQL statement to specify data to export when you're using the TransferText action or the TransferSpreadsheet action. Instead of using a SQL statement, you must first create a query and then specify the name of the query in the Table Name argument.

 _Comparisons Involving Null Values_

In Access versions 1.x and 2.0, if you compare two expressions within a macro condition by using a comparison operator and one of the expressions is  **Null**, Access Basic will return **True** or **False** for the comparison, depending on which comparison operator you use. In Access 2000 and later, Visual Basic returns **Null** for a comparison in which one expression is **Null**. To determine whether the comparison evaluates to **Null**, use the **IsNull** function to check the result of the comparison.


