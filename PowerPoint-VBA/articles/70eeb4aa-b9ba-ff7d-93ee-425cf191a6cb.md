
# AnimationBehavior Object (PowerPoint)

 **Last modified:** July 28, 2015

Represents the behavior of an animation effect, the main animation sequence, or an interactive animation sequence. The  **AnimationBehavior** object is a member of the ** [AnimationBehaviors](40e11093-5cbd-c8d3-04b5-4cd7de97bfa7.md)** collection.

## Example

Use  [Behaviors](e5335758-2f92-ccbc-a665-b6d5947e79f2.md)(index), where index is the number of the behavior in the sequence of behaviors, to return a single  **AnimationBehavior** object. The following example sets the positions of the a rotation's starting and ending points. This example assumes that the first behavior for the main animation sequence is a ** [RotationEffect](d0fc5520-dbbd-a44a-b811-51fd299c4587.md)**object.


```
Sub Change()

    With ActivePresentation.Slides(1).TimeLine.MainSequence(1) _

            .Behaviors(1).RotationEffect

        .From = 1

        .To = 180

    End With

End Sub
```


## See also


#### Concepts


 [PowerPoint Object Model Reference](00acd64a-5896-0459-39af-98df2849849e.md)
#### Other resources


 [AnimationBehavior Object Members](bf4580a3-3ad4-6158-8c72-2dcf9ded4202.md)
