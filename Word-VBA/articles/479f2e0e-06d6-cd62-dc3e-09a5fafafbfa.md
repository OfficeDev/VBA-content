
# Selection.TypeBackspace Method (Word)

Deletes the character preceding a collapsed selection (an insertion point).


## Syntax

 _expression_ . **TypeBackspace**

 _expression_ Required. A variable that represents a **[Selection](7b574a91-c33e-ecfd-6783-6b7528b2ed8f.md)** object.


## Remarks

This method corresponds to functionality of the BACKSPACE key. If the selection isn't collapsed to an insertion point, the selection is deleted.


## Example

This example deletes the character preceding the insertion point (the collapsed selection).


```vb
With Selection 
 .Collapse Direction:=wdCollapseEnd 
 .TypeBackspace 
End With
```

This example extends the selection to the end of the current paragraph (including the paragraph mark) and then deletes the selection.




```vb
With Selection 
 .EndOf Unit:=wdParagraph, Extend:=wdExtend 
 .TypeBackspace 
End With
```


## See also


#### Concepts


[Selection Object](7b574a91-c33e-ecfd-6783-6b7528b2ed8f.md)
#### Other resources


[Selection Object Members](71e67a43-d40a-ad9a-8ef2-c5c487733e0d.md)
