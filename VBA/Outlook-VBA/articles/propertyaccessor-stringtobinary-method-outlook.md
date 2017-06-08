---
title: PropertyAccessor.StringToBinary Method (Outlook)
keywords: vbaol11.chm1976
f1_keywords:
- vbaol11.chm1976
ms.prod: outlook
api_name:
- Outlook.PropertyAccessor.StringToBinary
ms.assetid: 1ea95601-a21f-47d2-7a3c-166c4984fc25
ms.date: 06/08/2017
---


# PropertyAccessor.StringToBinary Method (Outlook)

Converts a string specified by  _Value_ to an array of bytes.


## Syntax

 _expression_ . **StringToBinary**( **_Value_** )

 _expression_ A variable that represents a **PropertyAccessor** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Value_|Required| **String**|A hexadecimal string value that is to be converted to an array of bytes.|

### Return Value

A  **Variant** value that represents an array of bytes returned from the conversion.


## Remarks

For more information on type conversion when using the  **PropertyAccessor** object, see[Best Practices for Getting and Setting Properties](http://msdn.microsoft.com/library/ec087bf8-cfac-9b20-3cb2-3bd308c5c63d%28Office.15%29.aspx).


## See also


#### Concepts


[PropertyAccessor Object](propertyaccessor-object-outlook.md)

