---
title: Range.Sort Method (Excel)
keywords: vbaxl10.chm144200
f1_keywords:
- vbaxl10.chm144200
ms.prod: excel
api_name:
- Excel.Range.Sort
ms.assetid: ede52b2f-9025-fc83-9c16-f09c6b89c5c2
ms.date: 06/08/2017
---


# Range.Sort Method (Excel)

Sorts a range of values.


## Syntax

 _expression_. **Sort** (**_Key1_**, **_Order1_**, **_Key2_**, **_Type_**, **_Order2_**, **_Key3_**, **_Order3_**, **_Header_**, **_OrderCustom_**, **_MatchCase_**, **_Orientation_**, **_SortMethod_**, **_DataOption1_**, **_DataOption2_**, **_DataOption3_**)

 _expression_ A variable that represents a **Range** object.


### Parameters

|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Key1_|Optional| **Variant**|Specifies the first sort field, either as a range name (String) or **Range** object; determines the values to be sorted.|
| _Order1_|Optional| **XlSortOrder**|Determines the sort order for the values specified in _Key1_.|
| _Key2_|Optional| **Variant**|Second sort field; cannot be used when sorting a pivot table.|
| _Type_|Optional| **Variant**|Specifies which elements are to be sorted.|
| _Order2_|Optional| **XlSortOrder**|Determines the sort order for the values specified in  _Key2_.|
| _Key3_|Optional| **Variant**|Third sort field; cannot be used when sorting a pivot table.|
| _Order3_|Optional| **XlSortOrder**|Determines the sort order for the values specified in _Key3_.|
| _Header_|Optional| **XlYesNoGuess**|Specifies whether the first row contains header information. **xlNo** is the default value; specify **xlGuess** if you want Excel to attempt to determine the header.|
| _OrderCustom_|Optional| **Variant**|Specifies a one-based integer offset into the list of custom sort orders.|
| _MatchCase_|Optional| **Variant**|Set to **True** to perform a case-sensitive sort, **False** to perform non-case sensitive sort; cannot be used with pivot tables.|
| _Orientation_|Optional| **XlSortOrientation**|Specifies if the sort should be by row (default) or column. Set **xlSortColumns** value to	1	to sort by column. Set **xlSortRows** value to 2 to sort by row. (This is the default value.) |
| _SortMethod_|Optional| **XlSortMethod**|Specifies the sort method.|
| _DataOption1_|Optional| **XlSortDataOption**|Specifies how to sort text in the range specified in _Key1_; does not apply to pivot table sorting.|
| _DataOption2_|Optional| **XlSortDataOption**|Specifies how to sort text in the range specified in  _Key2_; does not apply to pivot table sorting.|
| _DataOption3_|Optional| **XlSortDataOption**|Specifies how to sort text in the range specified in _Key3_; does not apply to pivot table sorting.|

### Return value

Variant


## Example

**Sample code provided by:** Holy Macro! Books, [Holy Macro! It's 2,500 Excel VBA Examples](http://www.mrexcel.com/store/index.php?l=product_detail&;p=1)

This example gets the value of the color of a cell in column A using the **ColorIndex** property, and then uses that value to sort the range by color.

```vb
Sub ColorSort()
   'Set up your variables and turn off screen updating.
   Dim iCounter As Integer
   Application.ScreenUpdating = False
   
   'For each cell in column A, go through and place the color index value of the cell in column C.
   For iCounter = 2 To 55
      Cells(iCounter, 3) = _
         Cells(iCounter, 1).Interior.ColorIndex
   Next iCounter
   
   'Sort the rows based on the data in column C
   Range("C1") = "Index"
   Columns("A:C").Sort key1:=Range("C2"), _
      order1:=xlAscending, header:=xlYes
   
   'Clear out the temporary sorting value in column C, and turn screen updating back on.
   Columns(3).ClearContents
   Application.ScreenUpdating = True
End Sub
```


## About the contributor
<a name="AboutContributor"> </a>

*Holy Macro! Books* publishes entertaining books for people who use Office. See the complete catalog at MrExcel.com. 


## See also

#### Concepts

[Range Object](range-object-excel.md)

