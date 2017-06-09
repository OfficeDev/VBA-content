---
title: Series.ErrorBar Method (Word)
keywords: vbawd10.chm123732120
f1_keywords:
- vbawd10.chm123732120
ms.prod: word
api_name:
- Word.Series.ErrorBar
ms.assetid: 5deb5b19-50f1-2327-0b54-0952cd89756e
ms.date: 06/08/2017
---


# Series.ErrorBar Method (Word)

Applies error bars to the series.


## Syntax

 _expression_ . **ErrorBar**( **_Direction_** , **_Include_** , **_Type_** , **_Amount_** , **_MinusValues_** )

 _expression_ A variable that represents a **[Series](series-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Direction_|Required| **[XlErrorBarDirection](xlerrorbardirection-enumeration-word.md)**|One of the enumeration values that specifies the error bar direction.|
| _Include_|Required| **[XlErrorBarInclude](xlerrorbarinclude-enumeration-word.md)**|One of the enumeration values that specifies the error bar parts to include.|
| _Type_|Required| **[XlErrorBarType](xlerrorbartype-enumeration-word.md)**|One of the enumeration values that specifies the error bar type.|
| _Amount_|Optional| **Variant**|The error amount. Used for only the positive error amount when Type is  **xlErrorBarTypeCustom** .|
| _MinusValues_|Optional| **Variant**|The negative error amount when Type is  **xlErrorBarTypeCustom** .|

## Example

The following example applies standard error bars along the y-axis for series one of the first chart in the active document. The error bars are applied in the positive and negative directions. The example should be run on a 2-D line chart.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.SeriesCollection(1).ErrorBar _ 
 Direction:=xlY, Include:=xlErrorBarIncludeBoth, _ 
 Type:=xlErrorBarTypeStError 
 End If 
End With
```


## See also


#### Concepts


[Series Object](series-object-word.md)

