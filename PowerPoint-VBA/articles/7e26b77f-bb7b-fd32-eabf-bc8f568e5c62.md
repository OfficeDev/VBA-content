
# SlideShowView.GotoNamedShow Method (PowerPoint)

 **Last modified:** July 28, 2015

Switches to the specified custom, or named, slide show during another slide show. When the slide show advances from the current slide, the next slide displayed will be the next one in the specified custom slide show, not the next one in current slide show.

## Syntax

 _expression_. **GotoNamedShow**( **_SlideShowName_**)

 _expression_A variable that represents a  **SlideShowView** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|SlideShowName|Required| **String**|The name of the custom slide show to be switched to.|

## Example

This example redefines the slide show running in slide show window one to include only the slides in the custom slide show named "Quick Show."


```
SlideShowWindows(1).View.GotoNamedShow "Quick Show"
```


## See also


#### Concepts


 [SlideShowView Object](403b30ef-b12f-3a3c-e8d8-19189fd762fe.md)
#### Other resources


 [SlideShowView Object Members](fe2aacef-7324-4d07-55e9-0dffcdbb2a6c.md)
