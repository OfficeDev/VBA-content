---
title: Validation.Add Method (Excel)
keywords: vbaxl10.chm532073
f1_keywords:
- vbaxl10.chm532073
ms.prod: excel
api_name:
- Excel.Validation.Add
ms.assetid: e02c9d5e-dbb1-7d37-d112-0c60e7a7ff03
ms.date: 06/08/2017
---


# Validation.Add Method (Excel)

Adds data validation to the specified range.


## Syntax

 _expression_ . **Add**( **_Type_** , **_AlertStyle_** , **_Operator_** , **_Formula1_** , **_Formula2_** )

 _expression_ A variable that represents a **Validation** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Type_|Required| **XlDVType**|The validation type.|
| _AlertStyle_|Optional| **Variant**|The validation alert style. Can be one of the following  **[XlDVAlertStyle](xldvalertstyle-enumeration-excel.md)** constants: **xlValidAlertInformation** , **xlValidAlertStop** , or **xlValidAlertWarning** .|
| _Operator_|Optional| **Variant**|The data validation operator. Can be one of the following  **[XlFormatConditionOperator](xlformatconditionoperator-enumeration-excel.md)** constants: **xlBetween** , **xlEqual** , **xlGreater** , **xlGreaterEqual** , **xlLess** , **xlLessEqual** , **xlNotBetween** , or **xlNotEqual** .|
| _Formula1_|Optional| **Variant**|The first part of the data validation equation. Value must not exceed 255 characters.|
| _Formula2_|Optional| **Variant**|The second part of the data validation when  _Operator_ is **xlBetween** or **xlNotBetween** (otherwise, this argument is ignored).|

## Remarks

The  **Add** method requires different arguments, depending on the validation type, as shown in the following table.



|**Validation type**|**Arguments**|
|:-----|:-----|
| **xlValidateCustom**| **Formula1** is required, **Formula2** is ignored. **Formula1** must contain an expression that evaluates to **True** when data entry is valid and **False** when data entry is invalid.|
| **xlInputOnly**| **AlertStyle** , **Formula1** , or **Formula2** are used.|
| **xlValidateList**| **Formula1** is required, **Formula2** is ignored. **Formula1** must contain either a comma-delimited list of values or a worksheet reference to this list.|
| **xlValidateWholeNumber** , **xlValidateDate** , **xlValidateDecimal** , **xlValidateTextLength** , or **xlValidateTime**|One of either  **Formula1** or **Formula2** must be specified, or both may be specified.|

## Example

This example adds data validation to cell E5.


```vb
With Range("e5").Validation 
 .Add Type:=xlValidateWholeNumber, _ 
 AlertStyle:= xlValidAlertStop, _ 
 Operator:=xlBetween, Formula1:="5", Formula2:="10" 
 .InputTitle = "Integers" 
 .ErrorTitle = "Integers" 
 .InputMessage = "Enter an integer from five to ten" 
 .ErrorMessage = "You must enter a number from five to ten" 
End With
```


## See also


#### Concepts


[Validation Object](validation-object-excel.md)

