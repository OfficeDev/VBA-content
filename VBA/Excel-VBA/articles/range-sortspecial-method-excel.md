---
title: Range.SortSpecial Method (Excel)
keywords: vbaxl10.chm144201
f1_keywords:
- vbaxl10.chm144201
ms.prod: excel
api_name:
- Excel.Range.SortSpecial
ms.assetid: 706420cb-989a-1b48-b051-ca6e5fe45824
ms.date: 06/08/2017
---


# Range.SortSpecial Method (Excel)

Uses East Asian sorting methods to sort the range, a PivotTable report, or uses the method for the active region if the range contains only one cell. For example, Japanese sorts in the order of the Kana syllabary.


## Syntax

 _expression_ . **SortSpecial**( **_SortMethod_** , **_Key1_** , **_Order1_** , **_Type_** , **_Key2_** , **_Order2_** , **_Key3_** , **_Order3_** , **_Header_** , **_OrderCustom_** , **_MatchCase_** , **_Orientation_** , **_DataOption1_** , **_DataOption2_** , **_DataOption3_** )

 _expression_ A variable that represents a **Range** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SortMethod_|Optional| **[XlSortMethod](xlsortmethod-enumeration-excel.md)**|The type of sort. Some of these constants may not be available to you, depending on the language support (U.S. English, for example) that you?ve selected or installed.|
| _Key1_|Optional| **Variant**|The first sort field, as either text (a PivotTable field or range name) or a  **[Range](range-object-excel.md)** object ("Dept" or Cells(1, 1), for example).|
| _Order1_|Optional| **[XlSortOrder](xlsortorder-enumeration-excel.md)**|The sort order for the field or range specified in the  _Key1_ argument.|
| _Type_|Optional| **Variant**|Specifies which elements are to be sorted. Use this argument only when sorting PivotTable reports.|
| _Key2_|Optional| **Variant**|The second sort field, as either text (a PivotTable field or range name) or a  **Range** object. If you omit this argument, there?s no second sort field. Cannot be used when sorting PivotTable reports.|
| _Order2_|Optional| **XlSortOrder**|The sort order for the field or range specified in the  _Key2_ argument. Cannot be used when sorting PivotTable reports.|
| _Key3_|Optional| **Variant**|The third sort field, as either text (a range name) or a  **Range** object. If you omit this argument, there?s no third sort field. Cannot be used when sorting PivotTable reports.|
| _Order3_|Optional| **XlSortOrder**|The sort order for the field or range specified in the  _Key3_ argument. Cannot be used when sorting PivotTable reports.|
| _Header_|Optional| **[XlYesNoGuess](xlyesnoguess-enumeration-excel.md)**|Specifies whether or not the first row contains headers. Cannot be used when sorting PivotTable reports.|
| _OrderCustom_|Optional| **Variant**|This argument is a one-based integer offset to the list of custom sort orders. If you omit  _OrderCustom_, (normal sort order) is used.|
| _MatchCase_|Optional| **Variant**| **True** to do a case-sensitive sort; **False** to do a sort that?s not case sensitive. Cannot be used when sorting PivotTable reports.|
| _Orientation_|Optional| **[XlSortOrientation](xlsortorientation-enumeration-excel.md)**|The sort orientation.|
| _DataOption1_|Optional| **[XlSortDataOption](xlsortdataoption-enumeration-excel.md)**|Specifies how to sort text in  _Key1_. Cannot be used when sorting PivotTable reports.|
| _DataOption2_|Optional| **XlSortDataOption**|Specifies how to sort text in  _Key2_. Cannot be used when sorting PivotTable reports.|
| _DataOption3_|Optional| **XlSortDataOption**|Specifies how to sort text in  _Key3_. Cannot be used when sorting PivotTable reports.|

### Return Value

Variant


## Remarks

 If no arguments are defined with the **Sort** method, Microsoft Excel will sort the selection, chosen to be sorted, in ascending order.


## Example

This example sorts the range A1:A5 using Pin Yin (phonetic Chinese sort order for characters). In order to sort Chinese characters, this example assumes the user has Chinese language support for Microsoft Excel. Even without Chinese language support, Excel will default to sorting any numbers placed within the specified range for this example. This example assumes there is data contained in the range A1:A5.


```vb
Sub SpecialSort() 
 
 Application.Range("A1:A5").SortSpecial SortMethod:=xlPinYin 
 
End Sub
```


## See also


#### Concepts


[Range Object](range-object-excel.md)

