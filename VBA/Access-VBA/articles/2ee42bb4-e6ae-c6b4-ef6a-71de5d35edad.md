
# TextBox.AsianLineBreak Property (Access)

Returns or sets a  **Boolean** indicating whether line breaks in text boxes follow rules governing East Asian languages. **True** to control line breaks based on East Asian language rules. Read/write.


## Syntax

 _expression_. **AsianLineBreak**

 _expression_ A variable that represents a **TextBox** object.


## Remarks

Setting the  **AsianLineBreak** property to **True** moves any punctuation marks and closing parentheses at the beginning of a line to the end of the previous line, and moves opening parentheses at the end of a line to the beginning of the next line.


## Example

This example sets all the text boxes on the specified form to break lines according to East Asian language rules.


```vb
Dim ctlLoop As Control 
 
For Each ctlLoop In Forms(0).Controls 
 If ctlLoop.ControlType = acTextBox Then 
 ctlLoop.AsianLineBreak = True 
 End If 
Next ctlLoop
```


## See also


#### Concepts


[TextBox Object](d74fbe9a-0d40-7d28-956f-a2bfd0cfee45.md)
#### Other resources


[TextBox Object Members](bb55abbc-902e-fc2d-bdff-063c55426cd0.md)
