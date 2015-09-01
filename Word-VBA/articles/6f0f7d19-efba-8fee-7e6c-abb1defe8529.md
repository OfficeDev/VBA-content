
# Range.EmphasisMark Property (Word)

 **Last modified:** July 28, 2015

Returns or sets the emphasis mark for a character or designated character string. Read/write  **WdEmphasisMark**.

## Syntax

 _expression_. **EmphasisMark**

 _expression_Required. A variable that represents a  ** [Range](15a7a1c4-5f3f-5b6e-60e9-29688de3f274.md)** object.


## Example

This example sets the emphasis mark over the fourth word in the active document to a comma.


```
ActiveDocument.Words(4).EmphasisMark = wdEmphasisMarkOverComma
```


## See also


#### Concepts


 [Range Object](15a7a1c4-5f3f-5b6e-60e9-29688de3f274.md)
#### Other resources


 [Range Object Members](3c4a36d9-2a80-5aaf-827b-275a52bfa193.md)
