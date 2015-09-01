
# PageSetup.RightMargin Property (Word)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Returns or sets the distance (in points) between the right edge of the page and the right boundary of the body text. Read/write  **Single**.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **RightMargin**

 _expression_A variable that represents a  ** [PageSetup](1879d601-80ad-4fc0-1a87-92e999b59f88.md)** object.


## Remarks
<a name="sectionSection1"> </a>

If the  ** [MirrorMargins](ae7c53d9-7669-fb22-323f-2ad3984e2dfa.md)**property is set to  **True**, the  **RightMargin** property controls the setting for outside margins and the ** [LeftMargin](873d6cf2-da9f-5d88-314f-9820284a54ee.md)**property controls the setting for inside margins.


## Example
<a name="sectionSection2"> </a>

This example displays the right margin setting for the active document. The  ** [PointsToInches](e3d6ab40-3919-55e0-5829-603fca24c226.md)**method is used to convert the result to inches.


```
With ActiveDocument.PageSetup 
 Msgbox "The right margin is set to " _ 
 &amp; PointsToInches(.RightMargin) &amp; " inches." 
End With
```

This example sets the right margin for section two in the selection. The  ** [InchesToPoints](67a7e59c-bc61-be03-852d-05fadebef148.md)**method is used to convert inches to points.




```
Selection.Sections(2).PageSetup.RightMargin = InchesToPoints(1)
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [PageSetup Object](1879d601-80ad-4fc0-1a87-92e999b59f88.md)
#### Other resources


 [PageSetup Object Members](9ff8b896-933b-1a19-19d5-5e5d87aab1b5.md)
