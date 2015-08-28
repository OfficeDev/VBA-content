
# ColorEffect.To Property (PowerPoint)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Sets or returns a  **ColorFormat** object that represents the RGB color value of an animation behavior. Read/write.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **To**

 _expression_A variable that represents a  **ColorEffect** object.


### Return Value

ColorFormat


## Remarks
<a name="sectionSection1"> </a>

Use this property in conjunction with the  **From** property to transition from one color to another.

Do not confuse this property with the  **ToX**or  **ToY**properties of the  ** [ScaleEffect](cb7c296e-a9ea-4ed6-87e0-a5d603da4f9f.md)**and  ** [MotionEffect](77a34f68-8806-22b8-149f-c28e0457e7e9.md)**objects, which are only used for scaling or motion effects.


## Example
<a name="sectionSection2"> </a>

The following example adds a color effect and changes its color from a light bluish green to yellow.


```
Sub AddAndChangeColorEffect()

    Dim effBlinds As Effect

    Dim tmlTiming As TimeLine

    Dim shpRectangle As Shape

    Dim animColor As AnimationBehavior

    Dim clrEffect As ColorEffect



    Set shpRectangle = ActivePresentation.Slides(1).Shapes _

        .AddShape(Type:=msoShapeRectangle, Left:=100, _

        Top:=100, Width:=50, Height:=50)

    Set tmlTiming = ActivePresentation.Slides(1).TimeLine

    Set effBlinds = tmlTiming.MainSequence.AddEffect(Shape:=shpRectangle, _

        effectId:=msoAnimEffectBlinds)

    Set animColor = tmlTiming.MainSequence(1).Behaviors _

        .Add(Type:=msoAnimTypeColor)

    Set clrEffect = animColor.ColorEffect



    clrEffect.From.RGB = RGB(Red:=255, Green:=255, Blue:=0)

    clrEffect.To.RGB = RGB(Red:=0, Green:=255, Blue:=255)

End Sub
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [ColorEffect Object](c21ca9cd-3aaa-35f7-3d0e-89ca155322f2.md)
#### Other resources


 [ColorEffect Object Members](7b7317c7-5504-52f5-2437-990acc1b702d.md)
