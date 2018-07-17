---
title: PropertyAccessor.LocalTimeToUTC Method (Outlook)
keywords: vbaol11.chm1975
f1_keywords:
- vbaol11.chm1975
ms.prod: outlook
api_name:
- Outlook.PropertyAccessor.LocalTimeToUTC
ms.assetid: c19f60b2-441f-77b3-eb83-9cfd899e3a52
ms.date: 06/08/2017
---


# PropertyAccessor.LocalTimeToUTC Method (Outlook)

Converts a date-time value specified by  _Value_ from the local time format to Coordinated Universal Time (UTC) format.


## Syntax

 _expression_ . **LocalTimeToUTC**( **_Value_** )

 _expression_ A variable that represents a **PropertyAccessor** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Value_|Required| **Date**|The date-time value to be converted from local time to UTC.|

### Return Value

A  **Date** value that represents _Value_ after being converted from local time to UTC.


## Remarks

For more information on type conversion when using the  **PropertyAccessor** object, see[Best Practices for Getting and Setting Properties](http://msdn.microsoft.com/library/ec087bf8-cfac-9b20-3cb2-3bd308c5c63d%28Office.15%29.aspx).


## See also


#### Concepts


[PropertyAccessor Object](propertyaccessor-object-outlook.md)

