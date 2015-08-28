
# Workbook.Styles Property (Excel)

 **Last modified:** July 28, 2015

Returns a  ** [Styles](146effdc-e007-814d-b110-f7bd944fc15f.md)**collection that represents all the styles in the specified workbook. Read-only.

## Syntax

 _expression_. **Styles**

 _expression_A variable that represents a  **Workbook** object.


## Example

This example deletes the user-defined style "Stock Quote Style" from the active workbook.


```
ActiveWorkbook.Styles("Stock Quote Style").Delete
```


## See also


#### Concepts


 [Workbook Object](8c00aa60-c974-eed3-0812-3c9625eb0d4c.md)
#### Other resources


 [Workbook Object Members](dce102a3-25de-3ff4-2ce5-bc56e08baca7.md)
