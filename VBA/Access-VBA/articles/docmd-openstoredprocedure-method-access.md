---
title: DoCmd.OpenStoredProcedure Method (Access)
keywords: vbaac10.chm4651
f1_keywords:
- vbaac10.chm4651
ms.prod: access
api_name:
- Access.DoCmd.OpenStoredProcedure
ms.assetid: 90e229f9-072a-8d41-4c9b-363501770c8c
ms.date: 06/08/2017
---


# DoCmd.OpenStoredProcedure Method (Access)

The  **OpenStoredProcedure** method carries out the OpenStoredProcedure action in Visual Basic.


## Syntax

 _expression_. **OpenStoredProcedure**( ** _ProcedureName_**, ** _View_**, ** _DataMode_** )

 _expression_ A variable that represents a **DoCmd** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ProcedureName_|Required|**Variant**|A string expression that's the valid name of a stored procedure in the current database. If you execute Visual Basic code containing the  **OpenStoredProcedure** method in a library database, Microsoft Access looks for the stored procedure with this name first in the library database, then in the current database.|
| _View_|Optional|**AcView**|A  **[AcView](acview-enumeration-access.md)** constant that specifies the view in which the stored procedure will open. The default value is **acViewNormal**.|
| _DataMode_|Optional|**AcOpenDataMode**|A  **[AcOpenDataMode](acopendatamode-enumeration-access.md)** constant that specifies the data entry mode for the stored procedure. The default value is **acEdit**.|

## Remarks

In a Microsoft Access project, you can use the  **OpenStoredProcedure** method to open a stored procedure in Datasheet view, stored procedure Design view, or Print Preview. This method runs the named stored procedure when opened in Datasheet view. You can select the data entry mode for the stored procedure and restrict the records that the stored procedure displays.

If you don't want to display the system messages that normally appear when a stored procedure is run (indicating it's a stored procedure and showing how many records will be affected), you can use the  **SetWarnings** method to suppress the display of these messages.


## Example

The following example opens the Employees stored procedure in Design view.


```vb
DoCmd.OpenStoredProcedure "Employees", 1
```


## See also


#### Concepts


[DoCmd Object](docmd-object-access.md)

