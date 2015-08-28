
# PublishObject.RangeEnd Property (PowerPoint)

 **Last modified:** July 28, 2015

Returns or sets the number of the last slide in a range of slides you are publishing as a Web presentation. Read/write.

## Syntax

 _expression_. **RangeEnd**

 _expression_A variable that represents a  **PublishObject** object.


### Return Value

Integer


## Example

This example publishes slides three through five of the active presentation to HTML. It names the published presentation Mallard.htm.


```
With ActivePresentation.PublishObjects(1)

    .FileName = "C:\Test\Mallard.htm"

    .SourceType = ppPublishSlideRange

    .RangeStart = 3

    .RangeEnd = 5

    .Publish

End With
```


## See also


#### Concepts


 [PublishObject Object](9419bec4-d2a6-6a2c-6400-4e2e270ff603.md)
#### Other resources


 [PublishObject Object Members](a5cd1fb8-f916-ee2c-6114-165f2e5c3c23.md)
