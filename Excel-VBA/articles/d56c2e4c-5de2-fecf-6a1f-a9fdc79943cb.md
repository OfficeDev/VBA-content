
# Workbook.DeleteNumberFormat Method (Excel)

 **Last modified:** July 28, 2015

Deletes a custom number format from the workbook.

## Syntax

 _expression_. **DeleteNumberFormat**( **_NumberFormat_**)

 _expression_A variable that represents a  **Workbook** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|NumberFormat|Required| **String**|Names the number format to be deleted.|

## Example

This example deletes the number format "000-00-0000" from the active workbook.


```
ActiveWorkbook.DeleteNumberFormat("000-00-0000")
```


## See also


#### Concepts


 [Workbook Object](8c00aa60-c974-eed3-0812-3c9625eb0d4c.md)
#### Other resources


 [Workbook Object Members](dce102a3-25de-3ff4-2ce5-bc56e08baca7.md)
