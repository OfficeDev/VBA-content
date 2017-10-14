---
title: WorksheetFunction.Match Method (Excel)
keywords: vbaxl10.chm137114
f1_keywords:
- vbaxl10.chm137114
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Match
ms.assetid: 901cdd78-e8fc-f149-66ff-5887f7099c96
ms.date: 06/08/2017
---


# WorksheetFunction.Match Method (Excel)

Returns the relative position of an item in an array that matches a specified value in a specified order. Use MATCH instead of one of the LOOKUP functions when you need the position of an item in a range instead of the item itself.


## Syntax

 _expression_ . **Match**( **_Arg1_** , **_Arg2_** , **_Arg3_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Lookup_value - the value you use to find the value you want in a table.|
| _Arg2_|Required| **Variant**|Lookup_array - a contiguous range of cells containing possible lookup values. Lookup_array must be an array or an array reference.|
| _Arg3_|Optional| **Variant**|Match_type - the number -1, 0, or 1. Match_type specifies how Microsoft Excel matches lookup_value with values in lookup_array.|

### Return Value

Double


## Remarks




- Lookup_value is the value you want to match in lookup_array. For example, when you look up a number in a telephone book, you are using the person's name as the lookup value, but the telephone number is the value you want. 
    
- Lookup_value can be a value (number, text, or logical value) or a cell reference to a number, text, or logical value. 
    

- If match_type is 1, MATCH finds the largest value that is less than or equal to lookup_value. Lookup_array must be placed in ascending order: ...-2, -1, 0, 1, 2, ..., A-Z, FALSE, TRUE. 
    
- If match_type is 0, MATCH finds the first value that is exactly equal to lookup_value. Lookup_array can be in any order. 
    
- If match_type is -1, MATCH finds the smallest value that is greater than or equal to lookup_value. Lookup_array must be placed in descending order: TRUE, FALSE, Z-A, ...2, 1, 0, -1, -2, ..., and so on. 
    
- If match_type is omitted, it is assumed to be 1. 
    

- MATCH returns the position of the matched value within lookup_array, not the value itself. For example, MATCH("b",{"a","b","c"},0) returns 2, the relative position of "b" within the array {"a","b","c"}.
    
- MATCH does not distinguish between uppercase and lowercase letters when matching text values.
    
- If MATCH is unsuccessful in finding a match, it returns the #N/A error value.
    
- If match_type is 0 and lookup_value is text, you can use the wildcard characters, question mark (?) and asterisk (*), in lookup_value. A question mark matches any single character; an asterisk matches any sequence of characters. If you want to find an actual question mark or asterisk, type a tilde (~) before the character.
    

## Example

 **Sample code provided by:** Holy Macro! Books,[Holy Macro! It?s 2,500 Excel VBA Examples](http://www.mrexcel.com/store/index.php?l=product_detail&;p=1)

For each value in the first column of the first worksheet, this example searches through the entire workbook for a matching value. If the macro finds a matching value, it sets the original value on the first worksheet to be bold.




```vb
Sub HighlightMatches()
    Application.ScreenUpdating = False
    
    'Declare variables
    Dim var As Variant, iSheet As Integer, iRow As Long, iRowL As Long, bln As Boolean
       
       'Set up the count as the number of filled rows in the first column of Sheet1.
       iRowL = Cells(Rows.Count, 1).End(xlUp).Row
       
       'Cycle through all the cells in that column:
       For iRow = 1 To iRowL
          'For every cell that is not empty, search through the first column in each worksheet in the
          'workbook for a value that matches that cell value.

          If Not IsEmpty(Cells(iRow, 1)) Then
             For iSheet = ActiveSheet.Index + 1 To Worksheets.Count
                bln = False
                var = Application.Match(Cells(iRow, 1).Value, Worksheets(iSheet).Columns(1), 0)
                
                'If you find a matching value, indicate success by setting bln to true and exit the loop;
                'otherwise, continue searching until you reach the end of the workbook.
                If Not IsError(var) Then
                   bln = True
                   Exit For
                End If
             Next iSheet
          End If
          
          'If you do not find a matching value, do not bold the value in the original list;
          'if you do find a value, bold it.
          If bln = False Then
             Cells(iRow, 1).Font.Bold = False
             Else
             Cells(iRow, 1).Font.Bold = True
          End If
       Next iRow
    Application.ScreenUpdating = True
End Sub
```


## About the Contributor
<a name="AboutContributor"> </a>

Holy Macro! Books publishes entertaining books for people who use Microsoft Office. See the complete catalog at MrExcel.com. 


## See also
<a name="AboutContributor"> </a>


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

