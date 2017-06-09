---
title: DisplayFormat.Characters Property (Excel)
keywords: vbaxl10.chm893074
f1_keywords:
- vbaxl10.chm893074
ms.prod: excel
api_name:
- Excel.DisplayFormat.Characters
ms.assetid: 42e0518f-204d-c0cd-2401-dd1fb8f142e4
ms.date: 06/08/2017
---


# DisplayFormat.Characters Property (Excel)

Returns a  **[Characters](characters-object-excel.md)** object that represents a range of characters within the text of the associated **[Range](range-object-excel.md)** object as it is displayed in the current user interface. Read-only.


## Syntax

 _expression_ . **Characters**( **_Start_** , **_Length_** )

 _expression_ A variable that represents a **[DisplayFormat](displayformat-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Start_|Optional| **Variant**|The first character to be returned. If this argument is either 1 or omitted, this property returns a range of characters starting with the first character.|
| _Length_|Optional| **Variant**|The number of characters to be returned. If this argument is omitted, this property returns the remainder of the string (everything after the  _Start_ character).|

### Return Value

Characters


## Remarks

The  **Characters** object is not a collection.


## See also


#### Concepts


[DisplayFormat Object](displayformat-object-excel.md)

