---
title: VarType Function
keywords: vblr6.chm1009057
f1_keywords:
- vblr6.chm1009057
ms.prod: office
ms.assetid: 7422fba5-7ea9-1d91-fc0e-5694c352d2d0
ms.date: 06/08/2017
---


# VarType Function



Returns an  **Integer** indicating the subtype of a [variable](vbe-glossary.md).
 **Syntax**
 **VarType(**_varname_**)**
The required  _varname_[argument](vbe-glossary.md) is a [Variant](vbe-glossary.md) containing any variable except a variable of a [user-defined type](vbe-glossary.md).
 
 ## Return Values


|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
|**vbEmpty**|0|[Empty](vbe-glossary.md) (uninitialized)|
|**vbNull**|1|[Null](vbe-glossary.md) (no valid data)|
|**vbInteger**|2|Integer|
|**vbLong**|3|Long integer|
|**vbSingle**|4|Single-precision floating-point number|
|**vbDouble**|5|Double-precision floating-point number|
|**vbCurrency**|6|Currency value|
|**vbDate**|7|Date value|
|**vbString**|8|String|
|**vbObject**|9|Object|
|**vbError**|10|Error value|
|**vbBoolean**|11|Boolean value|
|**vbVariant**|12|**Variant** (used only with[arrays](vbe-glossary.md) of variants)|
|**vbDataObject**|13|A data access object|
|**vbDecimal**|14|Decimal value|
|**vbByte**|17|Byte value|
|**vbLongLong**|20|[LongLong](longlong-data-type.md) integer (Valid on 64-bit platforms only.)|
|**vbUserDefinedType**|36|Variants that contain user-defined types|
|**vbArray**|8192|Array|

 **Note**  These [constants](vbe-glossary.md) are specified by Visual Basic for Applications. The names can be used anywhere in your code in place of the actual values.

 **Remarks**
The  **VarType** function never returns the value for **vbArray** by itself. It is always added to some other value to indicate an array of a particular type. The constant **vbVariant** is only returned in conjunction with **vbArray** to indicate that the argument to the **VarType** function is an array of type **Variant**. For example, the value returned for an array of integers is calculated as **vbInteger** + **vbArray**, or 8194. If an object has a default [property](vbe-glossary.md),  **VarType** **(**_object_**)** returns the type of the object's default property.

## Example

This example uses the  **VarType** function to determine the subtype of a variable.


```vb
Dim IntVar, StrVar, DateVar, MyCheck
' Initialize variables.
IntVar = 459: StrVar = "Hello World": DateVar = #2/12/69# 
MyCheck = VarType(IntVar)    ' Returns 2.
MyCheck = VarType(DateVar)    ' Returns 7.
MyCheck = VarType(StrVar)    ' Returns 8.

```


