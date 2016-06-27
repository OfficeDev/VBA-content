
# Selection.ExtendMode Property (Word)

 **True** if Extend mode is active. Read/write **Boolean** .


## Syntax

 _expression_ . **ExtendMode**

 _expression_ An expression that returns a **[Selection](7b574a91-c33e-ecfd-6783-6b7528b2ed8f.md)** object.


## Remarks

When Extend mode is active, the Extend argument of the following methods is  **True** by default: **[EndKey](4f27681c-1117-99c2-1aba-bd97082bb8ba.md)** , **[HomeKey](24264193-d610-acbc-b393-de41fd55e976.md)** , **[MoveDown](d3ea31e8-04a5-c342-24ca-c93ac1a1258e.md)** , **[MoveLeft](23c22588-e774-f70f-28ea-81b1a54c0dd5.md)** , **[MoveRight](fcac96c7-7189-87b2-d800-9d161edb1e09.md)** , and **[MoveUp](46993371-c916-06b5-a644-960f8a283536.md)** . Also, the letters "EXT" appear on the status bar.

This property can only be set during run time; attempts to set it in Immediate mode are ignored. The Extend arguments of the  **[EndOf](33aa094b-17f9-3572-f66f-59692c57dc01.md)** and **[StartOf](570df152-3579-d7a6-f555-86c9da229e1b.md)** methods are not affected by this property.


## Example

This example moves to the beginning of the paragraph and selects the paragraph plus the next two sentences.


```vb
With Selection 
 .MoveUp Unit:=wdParagraph 
 .ExtendMode = True 
 .MoveDown Unit:=wdParagraph 
 .MoveRight Unit:=wdSentence, Count:=2 
End With
```

This example collapses the current selection, turns on Extend mode, and selects the current sentence.




```vb
With Selection 
 .Collapse 
 .ExtendMode = True 
 ' Select current word. 
 .Extend 
 ' Select current sentence. 
 .Extend 
End With
```


## See also


#### Concepts


[Selection Object](7b574a91-c33e-ecfd-6783-6b7528b2ed8f.md)
#### Other resources


[Selection Object Members](71e67a43-d40a-ad9a-8ef2-c5c487733e0d.md)
