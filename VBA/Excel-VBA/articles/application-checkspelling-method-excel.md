---
title: Application.CheckSpelling Method (Excel)
keywords: vbaxl10.chm133091
f1_keywords:
- vbaxl10.chm133091
ms.prod: excel
api_name:
- Excel.Application.CheckSpelling
ms.assetid: dfae0789-4635-5ec5-5146-c5a1acefa306
ms.date: 06/08/2017
---


# Application.CheckSpelling Method (Excel)

Checks the spelling of a single word.


## Syntax

 _expression_ . **CheckSpelling**( **_Word_** , **_CustomDictionary_** , **_IgnoreUppercase_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Word_|Required| **String**| (used only with **Application** object). The word you want to check.|
| _CustomDictionary_|Optional| **Variant**|A string that indicates the file name of the custom dictionary to be examined if the word isn't found in the main dictionary. If this argument is omitted, the currently specified dictionary is used.|
| _IgnoreUppercase_|Optional| **Variant**| **True** to have Microsoft Excel ignore words that are all uppercase. **False** to have Microsoft Excel check words that are all uppercase. If this argument is omitted, the current setting will be used.|

### Return Value

 **True** if the word is found in one of the dictionaries; otherwise **False** .


## Remarks

To check headers, footers, and objects on a worksheet, use this method on a  **[Worksheet](worksheet-object-excel.md)** object.

To check only cells and notes, use this method with the object returned by the  **[Cells](application-cells-property-excel.md)** property.


## See also


#### Concepts


[Application Object](application-object-excel.md)

