
# View.PrintOptions Property (PowerPoint)

 **Last modified:** July 28, 2015

Returns a  ** [PrintOptions](19ce56ba-b0d0-4086-db86-e32feade70bd.md)**object that represents print options that are saved with the specified presentation. Read-only.

## Syntax

 _expression_. **PrintOptions**

 _expression_A variable that represents a  **View** object.


### Return Value

PrintOptions


## Example

This example causes hidden slides in the active presentation to be printed, and it scales the printed slides to fit the paper size.


```
With Application.ActivePresentation

    With .PrintOptions

        .PrintHiddenSlides = True

        .FitToPage = True

    End With

    .PrintOut

End With
```


## See also


#### Concepts


 [View Object](333e8b59-398d-4575-d37b-bfb1d3503089.md)
#### Other resources


 [View Object Members](3330372c-8497-8cce-981b-3b64700eb915.md)
