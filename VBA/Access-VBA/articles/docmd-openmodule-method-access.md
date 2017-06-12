---
title: DoCmd.OpenModule Method (Access)
keywords: vbaac10.chm4161
f1_keywords:
- vbaac10.chm4161
ms.prod: access
api_name:
- Access.DoCmd.OpenModule
ms.assetid: 3d0b1599-6f52-e369-55e4-7fdc1c370953
ms.date: 06/08/2017
---


# DoCmd.OpenModule Method (Access)

The  **OpenModule** method carries out the OpenModule action in Visual Basic.


## Syntax

 _expression_. **OpenModule**( ** _ModuleName_**, ** _ProcedureName_** )

 _expression_ A variable that represents a **DoCmd** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ModuleName_|Optional|**Variant**| A string expression that's the valid name of the Visual Basic module you want to open. If you leave this argument blank, Microsoft Access searches all the standard modules in the database for the procedure you selected with the _procedurename_ argument and opens the module containing the procedure to that procedure. If you execute Visual Basic code containing the **OpenModule** method in a library database, Microsoft Access looks for the module with this name first in the library database, then in the current database.|
| _ProcedureName_|Optional|**Variant**|A string expression that's the valid name for the procedure you want to open the module to. If you leave this argument blank, the module opens to the Declarations section.|

## Remarks

You can use the  **OpenModule** method to open a specified Visual Basic module at a specified procedure. This can be a Sub procedure, a Function procedure, or an event procedure.

You must include at least one of the two OpenModule action arguments. If you enter a value for both arguments, Microsoft Access opens the specified module at the specified procedure.


## Example

The following example opens the Utility Functions module to the IsLoaded( )  **Function** procedure:


```vb
DoCmd.OpenModule "Utility Functions", "IsLoaded"
```


## See also


#### Concepts


[DoCmd Object](docmd-object-access.md)

