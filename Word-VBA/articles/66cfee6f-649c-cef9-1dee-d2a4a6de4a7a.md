
# Find.Frame Property (Word)

 **Last modified:** July 28, 2015

Returns a  ** [Frame](d36d3361-9e93-7dd9-b8c9-0ce503e03810.md)** object that represents the frame formatting for the specified style or find-and-replace operation. Read-only.

## Syntax

 _expression_. **Frame**

 _expression_A variable that represents a  ** [Find](da822788-cad5-992a-a835-18cc574cc324.md)** object.


## Example

This example finds the first frame with wrap around formatting. If such a frame is found, a message is displayed on the status bar.


```
With ActiveDocument.Content.Find 
 .Text = "" 
 .Frame.TextWrap = True 
 .Execute Forward:=True, Wrap:=wdFindContinue, Format:=True 
 If .Found = True Then StatusBar = "Frame was found" 
 .Parent.Select 
End With
```


## See also


#### Concepts


 [Find Object](da822788-cad5-992a-a835-18cc574cc324.md)
#### Other resources


 [Find Object Members](21f00da0-4c84-ace3-fc79-a55a9ed64360.md)
