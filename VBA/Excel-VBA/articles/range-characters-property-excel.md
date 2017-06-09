---
title: Range.Characters Property (Excel)
keywords: vbaxl10.chm144092
f1_keywords:
- vbaxl10.chm144092
ms.prod: excel
api_name:
- Excel.Range.Characters
ms.assetid: 5011b6d3-23ab-e2a8-9616-c4c73d3ae60e
ms.date: 06/08/2017
---


# Range.Characters Property (Excel)

Returns a  **[Characters](characters-object-excel.md)** object that represents a range of characters within the object text. You can use the **Characters** object to format characters within a text string.


## Syntax

 _expression_ . **Characters**( **_Start_** , **_Length_** )

 _expression_ A variable that represents a **Range** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Start_|Optional| **Variant**|The first character to be returned. If this argument is either 1 or omitted, this property returns a range of characters starting with the first character.|
| _Length_|Optional| **Variant**|The number of characters to be returned. If this argument is omitted, this property returns the remainder of the string (everything after the  _Start_ character).|

## Remarks

The  **Characters** object isn't a collection.


## Example

This example formats the third character in cell A1 on Sheet1 as bold.


```vb
With Worksheets("Sheet1").Range("A1") 
 .Value = "abcdefg" 
 .Characters(3, 1).Font.Bold = True 
End With
```


## See also


#### Concepts


[Range Object](range-object-excel.md)

