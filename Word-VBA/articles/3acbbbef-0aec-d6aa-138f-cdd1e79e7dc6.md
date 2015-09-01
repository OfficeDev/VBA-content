
# Global.LinesToPoints Method (Word)

 **Last modified:** July 28, 2015

Converts a measurement from lines to points (1 line = 12 points). Returns the converted measurement as a  **Single**.

## Syntax

 _expression_. **LinesToPoints**( **_Lines_**)

 _expression_A variable that represents a  ** [Global](b91e7459-08d5-ea8c-42e0-f7b9bfd1a72c.md)** object. Optional.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Lines|Required| **Single**|The line value to be converted to points.|

### Return Value

Single


## Example

This example sets the paragraph line spacing in the selection to three lines.


```
With Selection.ParagraphFormat 
 .LineSpacingRule = wdLineSpaceMultiple 
 .LineSpacing = LinesToPoints(3) 
End With
```


## See also


#### Concepts


 [Global Object](b91e7459-08d5-ea8c-42e0-f7b9bfd1a72c.md)
#### Other resources


 [Global Object Members](35050f7b-bc46-4795-ec17-f68e263c8af0.md)
