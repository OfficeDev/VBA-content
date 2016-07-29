
# Presentation.RemovePersonalInformation Property (PowerPoint)

Determines whether Microsoft PowerPoint should remove all user information from comments and revisions upon saving a presentation. Read/write.


## Syntax

 _expression_. **RemovePersonalInformation**

 _expression_ A variable that represents a **Presentation** object.


### Return Value

MsoTriState


## Remarks

The value of the  **RemovePersonalInformation** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**| Comments, revisions, and personal information remain in the presentation.|
|**msoTrue**| Removes comments, revisions, and personal information when saving presentation.|

## Example

This example sets the active presentation to remove personal information the next time the user saves it.


```vb
Sub RemovePersonalInfo()

    ActivePresentation.RemovePersonalInformation = msoTrue

End Sub
```


## See also


#### Concepts


[Presentation Object](ec75cf52-69f8-d35b-0a26-4a8da8a9683f.md)
#### Other resources


[Presentation Object Members](b3538c7e-5fd9-d34d-ab5c-0105dbd516d0.md)
