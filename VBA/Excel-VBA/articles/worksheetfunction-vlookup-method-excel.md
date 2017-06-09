---
title: WorksheetFunction.VLookup Method (Excel)
keywords: vbaxl10.chm137123
f1_keywords:
- vbaxl10.chm137123
ms.prod: excel
api_name:
- Excel.WorksheetFunction.VLookup
ms.assetid: 1b84b1f5-b557-3a5c-0787-7c19a9800580
ms.date: 06/08/2017
---


# WorksheetFunction.VLookup Method (Excel)

Searches for a value in the first column of a table array and returns a value in the same row from another column in the table array. 


## Syntax

 _expression_ . **VLookup**( **_Arg1_** , **_Arg2_** , **_Arg3_** , **_Arg4_** )

 _expression_ A variable that represents a **[WorksheetFunction](worksheetfunction-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Lookup_value - the value to search in the first column of the table array. Lookup_value can be a value or a reference. If lookup_value is smaller than the smallest value in the first column of table_array, VLOOKUP returns the #N/A error value.|
| _Arg2_|Required| **Variant**|Table_array - two or more columns of data. Use a reference to a range or a range name. The values in the first column of table_array are the values searched by lookup_value. These values can be text, numbers, or logical values. Uppercase and lowercase text are equivalent. |
| _Arg3_|Required| **Variant**|Col_index_num - the column number in table_array from which the matching value must be returned. A col_index_num of 1 returns the value in the first column in table_array; a col_index_num of 2 returns the value in the second column in table_array, and so on.|
| _Arg4_|Optional| **Variant**|Range_lookup - a logical value that specifies whether you want the  **VLookup** method to find an exact match or an approximate match:|

### Return Value

Variant


## Remarks

The V in  **VLookup** stands for vertical. Use the **VLookup** method instead of the **[HLookup](worksheetfunction-hlookup-method-excel.md)** method when your comparison values are located in a column to the left of the data that you want to find.


- If Col_index_num is less than 1, the  **VLookup** method generates an error.
    
- If Col_index_num is greater than the number of columns in table_array, the  **VLookup** method generates an error.
    

-  If Range_lookup is TRUE or omitted, an exact or approximate match is returned. If an exact match is not found, the next largest value that is less than lookup_value is returned. The values in the first column of table_array must be placed in ascending sort order; otherwise, the **VLookup** method may not give the correct value.
    
- If Range_lookup is FALSE, the  **VLookup** method will only find an exact match. In this case, the values in the first column of table_array do not need to be sorted. If there are two or more values in the first column of table_array that match the lookup_value, the first value found is used. If an exact match is not found, an error is generated.
    

- When searching text values in the first column of table_array, ensure that the data in the first column of table_array does not have leading spaces, trailing spaces, inconsistent use of straight ( ' or " ) and curly ( ? or ?) quotation marks, or nonprinting characters. In these cases, the  **VLookup** method may give an incorrect or unexpected value. For information about how to clean or trim values, see the **[Clean](worksheetfunction-clean-method-excel.md)** and **[Trim](worksheetfunction-trim-method-excel.md)** methods.
    
- When searching number or date values, ensure that the data in the first column of table_array is not stored as text values. In this case, the  **VLookup** method may give an incorrect or unexpected value.
    
- If range_lookup is FALSE and lookup_value is text, then you can use the wildcard characters, question mark (?) and asterisk (*), in lookup_value. A question mark matches any single character; an asterisk matches any sequence of characters. If you want to find an actual question mark or asterisk, type a tilde (~) preceding the character.
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

