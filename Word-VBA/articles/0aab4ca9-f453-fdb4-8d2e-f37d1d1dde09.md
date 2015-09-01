
# Options.PasteAdjustParagraphSpacing Property (Word)

 **Last modified:** July 28, 2015

 **True** if Microsoft Word automatically adjusts the spacing of paragraphs when cutting and pasting selections. Read/write **Boolean**.

## Syntax

 _expression_. **PasteAdjustParagraphSpacing**

 _expression_A variable that represents a  ** [Options](873b7b99-3fe1-fd89-9ece-a9355cb827dc.md)** object.


## Example

This example sets Word to automatically adjust the spacing of paragraphs when cutting and pasting selections if the option has been disabled.


```
Sub AdjustParaSpace() 
 With Options 
 If .PasteAdjustParagraphSpacing = False Then 
 .PasteAdjustParagraphSpacing = True 
 End If 
 End With 
End Sub
```


## See also


#### Concepts


 [Options Object](873b7b99-3fe1-fd89-9ece-a9355cb827dc.md)
#### Other resources


 [Options Object Members](76cd9dfe-6bbb-4c3d-0bfc-79a62bedd15e.md)
