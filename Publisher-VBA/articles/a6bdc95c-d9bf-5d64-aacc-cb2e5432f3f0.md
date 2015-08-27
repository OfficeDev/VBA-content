
# WebCheckBox.ReturnDataLabel Property (Publisher)

 **Last modified:** July 28, 2015

Returns or sets a  **String** that represents the text used by the Web page to label the specified Web object when the page is submitted. Read/write.

## Syntax

 _expression_. **ReturnDataLabel**

 _expression_A variable that represents a  **WebCheckBox** object.


## Example

This example creates a new Web text box and specifies the label for the text in the text box when the page is submitted.


```
Sub LabelWebTextBoxControl() 
 With ActiveDocument.Pages(1).Shapes _ 
 .AddWebControl(Type:=pbWebControlSingleLineTextBox, _ 
 Left:=100, Top:=100, Width:=300, Height:=15).WebTextBox 
 .DefaultText = "Please enter your name here" 
 .Limit = 70 
 .RequiredControl = msoTrue 
 .ReturnDataLabel = "Full_Name" 
 End With 
End Sub
```

