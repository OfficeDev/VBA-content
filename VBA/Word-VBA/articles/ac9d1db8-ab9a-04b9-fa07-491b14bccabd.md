
# Border.Color Property (Word)

Returns or sets the 24-bit color for the specified  **Border** object.


## Syntax

 _expression_ . **Color**

 _expression_ Required. A variable that represents a **[Border](be48c020-b86c-c004-ce1c-76d9edae9791.md)** object.


## Remarks

This property can be any valid  **WdColor** constant or a value returned by Visual Basic's **RGB** function.


## Example

This example adds a dotted indigo border around each cell in the first table.


```vb
If ActiveDocument.Tables.Count >= 1 Then 
 For Each aBorder In ActiveDocument.Tables(1).Borders 
 aBorder.Color = wdColorIndigo 
 aBorder.LineStyle = wdLineStyleDashDot 
 aBorder.LineWidth = wdLineWidth075pt 
 Next aBorder 
End If
```


## See also


#### Concepts


[Border Object](be48c020-b86c-c004-ce1c-76d9edae9791.md)
#### Other resources


[Border Object Members](0c2f320b-8f74-961b-297e-dc068db579aa.md)
