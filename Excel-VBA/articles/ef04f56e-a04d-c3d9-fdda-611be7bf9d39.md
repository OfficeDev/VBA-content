
# Workbook.Permission Property (Excel)

 **Last modified:** July 28, 2015

Returns a  **Permission** object that represents the permission settings in the specified workbook.

## Syntax

 _expression_. **Permission**

 _expression_A variable that represents a  **Workbook** object.


## Example

The following example returns the permission settings for the active workbook.


```
Dim objPermission As Permission 
 
Set objPermission = ActiveWorkbook.Permission
```


## See also


#### Concepts


 [Workbook Object](8c00aa60-c974-eed3-0812-3c9625eb0d4c.md)
#### Other resources


 [Workbook Object Members](dce102a3-25de-3ff4-2ce5-bc56e08baca7.md)
