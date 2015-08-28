
# Range.Consolidate Method (Excel)

 **Last modified:** July 28, 2015

Consolidates data from multiple ranges on multiple worksheets into a single range on a single worksheet.  **Variant**.

## Syntax

 _expression_. **Consolidate**( **_Sources_**,  **_Function_**,  **_TopRow_**,  **_LeftColumn_**,  **_CreateLinks_**)

 _expression_A variable that represents a  **Range** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Sources|Optional| **Variant**|The sources of the consolidation as an array of text reference strings in R1C1-style notation. The references must include the full path of sheets to be consolidated.|
|Function|Optional| **Variant**|One of the constants of  ** [XlConsolidationFunction](a3d0e4c0-8463-340c-a258-219d49f715d7.md)** which specifies the type of consolidation.|
|TopRow|Optional| **Variant**| **True** to consolidate data based on column titles in the top row of the consolidation ranges. **False** to consolidate data by position. The default value is **False**.|
|LeftColumn|Optional| **Variant**| **True** to consolidate data based on row titles in the left column of the consolidation ranges. **False** to consolidate data by position. The default value is **False**.|
|CreateLinks|Optional| **Variant**| **True** to have the consolidation use worksheet links. **False** to have the consolidation copy the data. The default value is **False**.|

### Return Value

Variant


## Example

This example consolidates data from Sheet2 and Sheet3 onto Sheet1, using the SUM function.


```
Worksheets("Sheet1").Range("A1").Consolidate _ 
 Sources:=Array("Sheet2!R1C1:R37C6", "Sheet3!R1C1:R37C6"), _ 
 Function:=xlSum
```


## See also


#### Concepts


 [Range Object](b8207778-0dcc-4570-1234-f130532cc8cd.md)
#### Other resources


 [Range Object Members](4336bf81-1e63-7e44-1792-baf366a027a7.md)
