
# NamedSlideShow Object (PowerPoint)

 **Last modified:** July 28, 2015

Represents a custom slide show, which is a named subset of slides in a presentation. 

## Remarks

The  **NamedSlideShow** object is a member of the ** [NamedSlideShows](9f20ff20-a81e-f771-5ef2-44b21ecfb055.md)**collection. The  **NamedSlideShows** collection contains all the named slide shows in the presentation.


## Example

Use  **NamedSlideShows**(index), where index is the custom slide show name or index number, to return a single  **NamedSlideShow** object. The following example deletes the custom slide show named "Quick Show."


```
ActivePresentation.SlideShowSettings _

    .NamedSlideShows("Quick Show").Delete
```

Use the  [SlideIDs](69c2a31e-bfb1-1a00-777f-4f5c46023ba0.md)property to return an array that contains the unique slide IDs for all the slides in the specified custom show. The following example displays the slide IDs for the slides in the custom slide show named "Quick Show."




```
idArray = ActivePresentation.SlideShowSettings _

    .NamedSlideShows("Quick Show").SlideIDs

For i = 1 To UBound(idArray)

    MsgBox idArray(i)

Next
```


## See also


#### Concepts


 [PowerPoint Object Model Reference](00acd64a-5896-0459-39af-98df2849849e.md)
#### Other resources


 [NamedSlideShow Object Members](a8ef0d6d-efe3-f63a-0e6f-68789aa58ebc.md)
