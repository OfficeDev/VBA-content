---
title: PropertyAccessor.UTCToLocalTime Method (Outlook)
keywords: vbaol11.chm1974
f1_keywords:
- vbaol11.chm1974
ms.prod: outlook
api_name:
- Outlook.PropertyAccessor.UTCToLocalTime
ms.assetid: a56311ac-60ac-4f51-5255-d6840bf6004d
ms.date: 06/08/2017
---


# PropertyAccessor.UTCToLocalTime Method (Outlook)

Converts the date-time value that is specified by  _Value_ and expressed in Coordinated Universal Time (UTC) to a value in local time.


## Syntax

 _expression_ . **UTCToLocalTime**( **_Value_** )

 _expression_ A variable that represents a **PropertyAccessor** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Value_|Required| **Date**|The date-time value to be converted from UTC to local time.|

### Return Value

A  **Date** value that represents _Value_ after being converted from UTC to local time.


## Remarks

For more information on type conversion when using the  **[PropertyAccessor](propertyaccessor-object-outlook.md)** object, see[Best Practices for Getting and Setting Properties](http://msdn.microsoft.com/library/ec087bf8-cfac-9b20-3cb2-3bd308c5c63d%28Office.15%29.aspx).


## See also


#### Concepts


[PropertyAccessor Object](propertyaccessor-object-outlook.md)

