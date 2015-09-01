
# TextFrame.HasText Property (Word)

 **Last modified:** July 28, 2015

 **True** if the specified shape has text associated with it. Read-only **Boolean**.

## Syntax

 _expression_. **HasText**

 _expression_A variable that represents a  ** [TextFrame](46f7e410-80d9-9fe9-2224-488b623f8592.md)** object.


## Example

If the second shape on the active document contains text, this example displays a message if the text overflows its frame.


```
Dim docActive As Document 
 
Set docActive = ActiveDocument 
With docActive.Shapes(2).TextFrame 
 If .HasText = True Then 
 If .Overflowing = True Then 
 Msgbox "Text overflows the frame." 
 End If 
 End If 
End With
```


## See also


#### Concepts


 [TextFrame Object](46f7e410-80d9-9fe9-2224-488b623f8592.md)
#### Other resources


 [TextFrame Object Members](bb2efcc6-474f-3de5-6d20-940be7549112.md)
