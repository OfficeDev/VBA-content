---
title: Range.Information Property (Word)
keywords: vbawd10.chm157155641
f1_keywords:
- vbawd10.chm157155641
ms.prod: word
api_name:
- Word.Range.Information
ms.assetid: 967e9a22-5f98-e4bd-557c-7367cb7c5d2b
ms.date: 06/08/2017
---


# Range.Information Property (Word)

Returns information about the specified range. Read-only  **Variant** .


## Syntax

 _expression_ . **Information**( **_Type_** )

 _expression_ Required. A variable that represents a **[Range](range-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Type_|Required| **WdInformation**|The information type.|

## Example

If the tenth word is in a table, this example selects the table.


```vb
If ActiveDocument.Words(10).Information(wdWithInTable) Then _ 
 ActiveDocument.Words(10).Tables(1).Select
```


## See also


#### Concepts


[Range Object](range-object-word.md)

