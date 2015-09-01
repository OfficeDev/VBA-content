
# Dialog.Execute Method (Word)

 **Last modified:** July 28, 2015

Applies the current settings of a Microsoft Word dialog box.

## Syntax

 _expression_. **Execute**

 _expression_Required. A variable that represents a  ** [Dialog](f90f6e6d-aaa0-c127-ab37-ca074144eff1.md)** object.


## Example

The following example enables the  **Keep with next** check box on the **Line and Page Breaks** tab in the **Paragraph** dialog box.


```
With Dialogs(wdDialogFormatParagraph) 
 .KeepWithNext = 1 
 .Execute 
End With
```


## See also


#### Concepts


 [Dialog Object](f90f6e6d-aaa0-c127-ab37-ca074144eff1.md)
#### Other resources


 [Dialog Object Members](f5c755d5-9fdf-bfb4-2c4b-8999ae176635.md)
