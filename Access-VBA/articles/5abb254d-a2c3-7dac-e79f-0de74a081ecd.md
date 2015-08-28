
# Workbook.NewSheet Event (Excel)

 **Last modified:** July 28, 2015

Occurs when a new sheet is created in the workbook.

## Syntax

 _expression_. **NewSheet**( **_Sh_**)

 _expression_A variable that represents a  **Workbook** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Sh|Required| **Object**|The new sheet. Can be a  ** [Worksheet](182b705e-854a-81cc-a4b0-59b942de55ae.md)** or ** [Chart](179c32ce-49bd-6f36-ea12-89fb5443f3ea.md)** object.|

### Return Value

Nothing


## Example

This example moves new sheets to the end of the workbook.


```
Private Sub Workbook_NewSheet(ByVal Sh as Object) 
 Sh.Move After:= Sheets(Sheets.Count) 
End Sub
```


## See also


#### Concepts


 [Workbook Object](8c00aa60-c974-eed3-0812-3c9625eb0d4c.md)
#### Other resources


 [Workbook Object Members](dce102a3-25de-3ff4-2ce5-bc56e08baca7.md)
