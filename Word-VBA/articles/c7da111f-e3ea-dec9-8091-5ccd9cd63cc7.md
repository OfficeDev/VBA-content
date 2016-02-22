
# Find.MatchByte Property (Word)

 **True** if Microsoft Word distinguishes between full-width and half-width letters or characters during a search. Read/write **Boolean** .


## Syntax

 _expression_ . **MatchByte**

 _expression_ A variable that represents a **[Find](da822788-cad5-992a-a835-18cc574cc324.md)** object.


## Example

This example searches for the term "マイクロソフト" in the specified range without distinguishing between full-width and half-width characters.


```vb
With Selection.Find 
    .ClearFormatting 
    .MatchWholeWord = True 
    .MatchByte = False 
    .Execute FindText:="マイクロソフト" 
End With
```


## See also


#### Concepts


[Find Object](da822788-cad5-992a-a835-18cc574cc324.md)
#### Other resources


[Find Object Members](21f00da0-4c84-ace3-fc79-a55a9ed64360.md)
