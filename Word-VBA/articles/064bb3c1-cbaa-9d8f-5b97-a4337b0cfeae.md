
# Selection.FootnoteOptions Property (Word)

 **Last modified:** July 28, 2015

Returns  ** [FootnoteOptions](5fdeb6d6-ce33-44f5-62c1-743fc3770457.md)**object that represents the footnotes in a selection.

## Syntax

 _expression_. **FootnoteOptions**

 _expression_A variable that represents a  ** [Selection](7b574a91-c33e-ecfd-6783-6b7528b2ed8f.md)** object.


## Example

This example sets the numbering rule in the selection to restart at the beginning of the new section.


```
Sub SetFootnoteOptionsRange() 
 Selection.FootnoteOptions.NumberingRule = wdRestartSection 
End Sub
```


## See also


#### Concepts


 [Selection Object](7b574a91-c33e-ecfd-6783-6b7528b2ed8f.md)
#### Other resources


 [Selection Object Members](71e67a43-d40a-ad9a-8ef2-c5c487733e0d.md)
