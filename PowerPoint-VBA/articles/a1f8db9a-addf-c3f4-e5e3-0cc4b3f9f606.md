
# AnimationBehavior.ColorEffect Property (PowerPoint)

 **Last modified:** July 28, 2015

Returns a  ** [ColorEffect](c21ca9cd-3aaa-35f7-3d0e-89ca155322f2.md)** object that represents the color properties for a specified animation behavior.

## Syntax

 _expression_. **ColorEffect**

 _expression_A variable that represents an  **AnimationBehavior** object.


### Return Value

ColorEffect


## Example

This example adds a shape to the first slide of the active presentation and sets a color effect behavior to change the fill color of the new shape.


```
Sub ChangeColorEffect()

    Dim sldFirst As Slide

    Dim shpHeart As Shape

    Dim effNew As Effect

    Dim bhvEffect As AnimationBehavior



    Set sldFirst = ActivePresentation.Slides(1)

    Set shpHeart = sldFirst.Shapes.AddShape(Type:=msoShapeHeart, _

        Left:=100, Top:=100, Width:=100, Height:=100)

    Set effNew = sldFirst.TimeLine.MainSequence.AddEffect _

        (Shape:=shpHeart, EffectID:=msoAnimEffectChangeFillColor, _

        Trigger:=msoAnimTriggerAfterPrevious)

    Set bhvEffect = effNew.Behaviors.Add(Type:=msoAnimTypeColor)



    With bhvEffect.ColorEffect

        .From.RGB = RGB(Red:=255, Green:=0, Blue:=0)

        .To.RGB = RGB(Red:=0, Green:=0, Blue:=255)

    End With

End Sub
```


## See also


#### Concepts


 [AnimationBehavior Object](70eeb4aa-b9ba-ff7d-93ee-425cf191a6cb.md)
#### Other resources


 [AnimationBehavior Object Members](bf4580a3-3ad4-6158-8c72-2dcf9ded4202.md)
