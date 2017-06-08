---
title: Validation.Modify Method (Excel)
keywords: vbaxl10.chm532085
f1_keywords:
- vbaxl10.chm532085
ms.prod: excel
api_name:
- Excel.Validation.Modify
ms.assetid: 4f6b435a-6ca6-8953-1bde-549b0bdc1774
ms.date: 06/08/2017
---


# Validation.Modify Method (Excel)

Modifies data validation for a range.


## Syntax

 _expression_ . **Modify**( **_Type_** , **_AlertStyle_** , **_Operator_** , **_Formula1_** , **_Formula2_** )

 _expression_ A variable that represents a **Validation** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Type_|Optional| **Variant**|An  **XlDVType** value that represents the validation type.|
| _AlertStyle_|Optional| **Variant**|An  **[XlDVAlertStyle](xldvalertstyle-enumeration-excel.md)** value that represents the validation alert style.|
| _Operator_|Optional| **Variant**|An  **[XlFormatConditionOperator](xlformatconditionoperator-enumeration-excel.md)** value that represents the data validation operator.|
| _Formula1_|Optional| **Variant**|The first part of the data validation equation.|
| _Formula2_|Optional| **Variant**|The second part of the data validation when  **Operator** is **xlBetween** or **xlNotBetween** ; otherwise, this argument is ignored.|

## Remarks

The  **Modify** method requires different arguments, depending on the validation type, as shown in the following table.



|**Validation type**|**Arguments**|
|:-----|:-----|
| **xlInputOnly**| **AlertStyle** , **Formula1** , and **Formula2** are not used.|
| **xlValidateCustom**| **Formula1** is required; **Formula2** is ignored. **Formula1** must contain an expression that evaluates to **True** when data entry is valid and **False** when data entry is invalid.|
| **xlValidateList**| **Formula1** is required; **Formula2** is ignored. **Formula1** must contain either a comma-delimited list of values or a worksheet reference to the list.|
| **xlValidateDate** , **xlValidateDecimal** , **xlValidateTextLength** , **xlValidateTime** , or **xlValidateWholeNumber**| **Formula1** or **Formula2** , or both, must be specified.|

## Example

This example changes data validation for cell E5.


```vb
Range("e5").Validation _ 
 .Modify xlValidateList, xlValidAlertStop, _ 
 xlBetween, "=$A$1:$A$10"
```


## See also


#### Concepts


[Validation Object](validation-object-excel.md)

