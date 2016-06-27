
# Find.Font Property (Word)

Returns or sets a  **[Font](bc97f4df-fc81-d6c8-e99a-d50dc793b7ae.md)** object that represents the character formatting of the specified object. Read/write **Font** .


## Syntax

 _expression_ . **Font**

 _expression_ A variable that represents a **[Find](da822788-cad5-992a-a835-18cc574cc324.md)** object.


## Remarks

To set this property, specify an expression that returns a  **[Font](bc97f4df-fc81-d6c8-e99a-d50dc793b7ae.md)** object.


## Example

This example finds the next range of text that's formatted with the Times New Roman font.


```vb
With Selection.Find 
 .ClearFormatting 
 .Font.Name = "Times New Roman" 
 .Execute FindText:="", ReplaceWith:="", Format:=True, _ 
 Forward:=True 
End With
```


## See also


#### Concepts


[Find Object](da822788-cad5-992a-a835-18cc574cc324.md)
#### Other resources


[Find Object Members](21f00da0-4c84-ace3-fc79-a55a9ed64360.md)
