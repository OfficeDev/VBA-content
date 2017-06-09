---
title: Application.CommandBars Property (Excel)
keywords: vbaxl10.chm183088
f1_keywords:
- vbaxl10.chm183088
ms.prod: excel
api_name:
- Excel.Application.CommandBars
ms.assetid: b1884d43-557b-47be-1cef-20404069b576
ms.date: 06/08/2017
---


# Application.CommandBars Property (Excel)

Returns a  **[CommandBars](http://msdn.microsoft.com/library/0e312e21-14ee-5055-d604-b66e61c53b47%28Office.15%29.aspx)** object that represents the Microsoft Excel command bars. Read-only.


## Syntax

 _expression_ . **CommandBars**

 _expression_ An expression that returns a **Application** object.


## Remarks

Used with the  **[Application](application-object-excel.md)** object, this property returns the set of built-in and custom command bars available to the application.

When a workbook is embedded in another application and activated by the user by double-clicking the workbook, using this property with a  **[Workbook](workbook-object-excel.md)** object returns the set of Microsoft Excel command bars available within the other application. At all other times, using this property with a **Workbook** object returns **Nothing** .

There is no programmatic way to return the set of command bars attached to a workbook.


## Example

This example deletes all custom command bars that aren't visible.


```vb
For Each bar In Application.CommandBars 
    If Not bar.BuiltIn And Not bar.Visible Then bar.Delete 
Next
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

