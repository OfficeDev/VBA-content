
# Application.Watches Property (Excel)

 **Last modified:** July 28, 2015

Returns a  ** [Watches](de403bcc-b927-90f6-75d7-9c936c7f58f7.md)**object representing a range which is tracked when the worksheet is recalculated.

## Syntax

 _expression_. **Watches**

 _expression_A variable that represents an  **Application** object.


## Example

This example creates a summation formula in cell A3, and then adds this cell to the Watch Window.


```
Sub AddWatch() 
 With Application 
 .Range("A1").Formula = 1 
 .Range("A2").Formula = 2 
 .Range("A3").Formula = "=Sum(A1:A2)" 
 .Range("A3").Select 
 .Watches.Add Source:=ActiveCell 
 End With 
End Sub
```


## See also


#### Concepts


 [Application Object](19b73597-5cf9-4f56-8227-b5211f657f6f.md)
#### Other resources


 [Application Object Members](4cb9ca42-8d07-cc9c-2d80-4eb9a5921e1e.md)
