
# Find.MatchByte Property (Word)

 **Last modified:** July 28, 2015

 **True** if Microsoft Word distinguishes between full-width and half-width letters or characters during a search. Read/write **Boolean**.

## Syntax

 _expression_. **MatchByte**

 _expression_A variable that represents a  ** [Find](da822788-cad5-992a-a835-18cc574cc324.md)** object.


## Example

This example searches for the term "ãƒžã‚¤ã‚¯ãƒ­ã‚½ãƒ•ãƒˆ" in the specified range without distinguishing between full-width and half-width characters.


```
With Selection.Find 
    .ClearFormatting 
    .MatchWholeWord = True 
    .MatchByte = False 
    .Execute FindText:="ãƒžã‚¤ã‚¯ãƒ­ã‚½ãƒ•ãƒˆ" 
End With
```


## See also


#### Concepts


 [Find Object](da822788-cad5-992a-a835-18cc574cc324.md)
#### Other resources


 [Find Object Members](21f00da0-4c84-ace3-fc79-a55a9ed64360.md)
