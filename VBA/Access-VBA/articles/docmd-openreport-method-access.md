---
title: DoCmd.OpenReport Method (Access)
keywords: vbaac10.chm4163
f1_keywords:
- vbaac10.chm4163
ms.prod: access
api_name:
- Access.DoCmd.OpenReport
ms.assetid: 3c08755a-5116-f085-d498-725dc12e62f1
ms.date: 06/08/2017
---


# DoCmd.OpenReport Method (Access)

The  **OpenReport** method carries out the OpenReport action in Visual Basic.


## Syntax

 _expression_. **OpenReport**( ** _ReportName_**, ** _View_**, ** _FilterName_**, ** _WhereCondition_**, ** _WindowMode_**, ** _OpenArgs_** )

 _expression_ A variable that represents a **DoCmd** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ReportName_|Required|**Variant**|A string expression that's the valid name of a report in the current database. If you execute Visual Basic code containing the  **OpenReport** method in a library database, Microsoft Access looks for the report with this name, first in the library database, then in the current database.|
| _View_|Optional|**AcView**|A  **[AcView](acview-enumeration-access.md)** constant that specifies the view in which the report will open. The default value is **acViewNormal**.|
| _FilterName_|Optional|**Variant**|A string expression that's the valid name of a query in the current database.|
| _WhereCondition_|Optional|**Variant**|A string expression that's a valid SQL WHERE clause without the word WHERE.|
| _WindowMode_|Optional|**AcWindowMode**|A  **[AcWindowMode](acwindowmode-enumeration-access.md)** constant that specifies the mode in which the form opens. The default valus is **acWindowNormal**.|
| _OpenArgs_|Optional|**Variant**|Sets the  **OpenArgs** property.|

## Remarks

You can use the  **OpenReport** method to open a report in Design view or Print Preview, or to print the report immediately. You can also restrict the records that are printed in the report.

The maximum length of the  _WhereCondition_ argument is 32,768 characters (unlike the Where Condition action argument in the Macro window, whose maximum length is 256 characters).


## Example

The following example prints Sales Report while using the existing query Report Filter.


```vb
DoCmd.OpenReport "Sales Report", acViewNormal, "Report Filter"
```


## See also


#### Concepts


[DoCmd Object](docmd-object-access.md)

