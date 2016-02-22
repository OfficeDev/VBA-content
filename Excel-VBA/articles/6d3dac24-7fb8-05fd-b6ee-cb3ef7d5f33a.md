
# Application.CalculateFullRebuild Method (Excel)

For all open workbooks, forces a full calculation of the data and rebuilds the dependencies.


## Syntax

 _expression_ . **CalculateFullRebuild**

 _expression_ A variable that represents an **Application** object.


## Remarks

Dependencies are the formulas that depend on other cells. For example, the formula "=A1" depends on cell A1. The  **CalculateFullRebuild** method is similar to re-entering all formulas.


## Example

This example compares the version of Microsoft Excel with the version of Excel in which the workbook was last calculated. If the two version numbers are different, a full calculation of the data in all open workbooks is performed and the dependencies are rebuilt.


```vb
Sub UseCalculateFullRebuild() 
 
 If Application.CalculationVersion <> _ 
 Workbooks(1).CalculationVersion Then 
 Application.CalculateFullRebuild 
 End If 
 
End Sub
```


## See also


#### Concepts


[Application Object](19b73597-5cf9-4f56-8227-b5211f657f6f.md)
#### Other resources


[Application Object Members](4cb9ca42-8d07-cc9c-2d80-4eb9a5921e1e.md)
