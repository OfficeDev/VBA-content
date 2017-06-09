---
title: DoCmd.OpenFunction Method (Access)
keywords: vbaac10.chm5161
f1_keywords:
- vbaac10.chm5161
ms.prod: access
api_name:
- Access.DoCmd.OpenFunction
ms.assetid: 56168394-9e83-f620-8b5e-680e824ec941
ms.date: 06/08/2017
---


# DoCmd.OpenFunction Method (Access)

Opens a user-defined function in a Microsoft SQL Server database for viewing in Microsoft Access.


## Syntax

 _expression_. **OpenFunction**( ** _FunctionName_**, ** _View_**, ** _DataMode_** )

 _expression_ A variable that represents a **DoCmd** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FunctionName_|Required|**Variant**|The name of the function to open.|
| _View_|Optional|**AcView**|A  **[AcView](acview-enumeration-access.md)** constant that specifies the view in which to open the function. The default value is **acViewNormal**.|
| _DataMode_|Optional|**AcOpenDataMode**|A  **[AcOpenDataMode](acopendatamode-enumeration-access.md)** constant that specifies the mode in which to open the function. The default value is **acEdit**.|

## Remarks

Use the  **AllFunctions** collection to retrieve information about the available user-defined functions in a SQL Server database.


## Example

The following example opens the first user-defined function in the current database in Design View and read-only mode.


```vb
Dim objFunction As AccessObject 
Dim strFunction As String 
 
Set objFunction = Application.AllFunctions(0) 
 
DoCmd.OpenFunction FunctionName:=objFunction.Name, _ 
 View:=acViewDesign, Mode:=acReadOnly 

```


## See also


#### Concepts


[DoCmd Object](docmd-object-access.md)

