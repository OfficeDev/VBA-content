
# TextRange.MajorityFont Property (Publisher)

 **Last modified:** July 28, 2015

Returns a  ** [Font](992fda94-2820-d665-0d78-efd4b5434731.md)** object that represents the font name most in use in a text range.

## Syntax

 _expression_. **MajorityFont**

 _expression_A variable that represents a  **TextRange** object.


### Return Value

Font


## Example

This example creates a new text box, fills it with text, checks if the font most in use is Tahoma, and if it isn't, changes the font to Tahoma.


```
Sub SetFontName() 
 Dim intCount As Integer 
 With ActiveDocument.Pages(1).Shapes _ 
 .AddTextbox(Orientation:=pbTextOrientationHorizontal, _ 
 Left:=100, Top:=100, Width:=100, Height:=100) _ 
 .TextFrame.TextRange 
 For intCount = 1 To 10 
 .InsertAfter NewText:="This is a test. " 
 Next intCount 
 If .MajorityFont <> "Tahoma" Then _ 
 .Font.Name = "Tahoma" 
 End With 
End Sub
```

