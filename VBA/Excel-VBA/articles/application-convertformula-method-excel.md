---
title: Application.ConvertFormula Method (Excel)
keywords: vbaxl10.chm133097
f1_keywords:
- vbaxl10.chm133097
ms.prod: excel
api_name:
- Excel.Application.ConvertFormula
ms.assetid: 6ed0a76c-9db5-f6ab-a91d-d4e1b6674c53
ms.date: 06/08/2017
---


# Application.ConvertFormula Method (Excel)

Converts cell references in a formula between the A1 and R1C1 reference styles, between relative and absolute references, or both.  **Variant** .


## Syntax

 _expression_ . **ConvertFormula**( **_Formula_** , **_FromReferenceStyle_** , **_ToReferenceStyle_** , **_ToAbsolute_** , **_RelativeTo_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Formula_|Required| **Variant**|A string that containis the formula you want to convert. This must be a valid formula, and it must begin with an equal sign.|
| _FromReferenceStyle_|Required| **[XlReferenceStyle](xlreferencestyle-enumeration-excel.md)**|The reference style of the formula.|
| _ToReferenceStyle_|Optional| **Variant**|A constant of  **XlReferenceStyle** specifying the reference style you want returned. If this argument is omitted, the reference style isn't changed; the formula stays in the style specified by _FromReferenceStyle_.|
| _ToAbsolute_|Optional| **Variant**|A constant of  **[XlReferenceType](xlreferencetype-enumeration-excel.md)** which specifies the converted reference type. If this argument is omitted, the reference type isn't changed.|
| _RelativeTo_|Optional| **Variant**|A  **Range** object that contains one cell. Relative references relate to this cell.|

### Return Value

Variant


## Remarks

There is a 255 character limit for the formula.


## Example

This example converts a SUM formula that contains R1C1-style references to an equivalent formula that contains A1-style references, and then it displays the result.


```
inputFormula = "=SUM(R10C2:R15C2)" 
MsgBox Application.ConvertFormula( _ 
 formula:=inputFormula, _ 
 fromReferenceStyle:=xlR1C1, _ 
 toReferenceStyle:=xlA1)
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

