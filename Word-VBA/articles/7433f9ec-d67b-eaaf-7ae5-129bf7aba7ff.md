
# Options.SnapToShapes Property (Word)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


 **True** if Word automatically aligns AutoShapes or East Asian characters with invisible gridlines that go through the vertical and horizontal edges of other AutoShapes or East Asian characters. Read/write **Boolean**.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **SnapToShapes**

 _expression_A variable that represents an  ** [Options](873b7b99-3fe1-fd89-9ece-a9355cb827dc.md)** collection.


## Remarks
<a name="sectionSection1"> </a>

This property creates additional invisible gridlines for each AutoShape.  **SnapToShapes** works independently of the **SnapToGrid** property.


## Example
<a name="sectionSection2"> </a>

This example sets Word to automatically align AutoShapes with invisible gridlines that go through the vertical and horizontal edges of other AutoShapes in a new document.


```
Options.SnapToShapes = True 
Documents.Add
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [Options Object](873b7b99-3fe1-fd89-9ece-a9355cb827dc.md)
#### Other resources


 [Options Object Members](76cd9dfe-6bbb-4c3d-0bfc-79a62bedd15e.md)
