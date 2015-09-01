
# Document.Frames Property (Word)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Returns a  ** [Frames](d0f526b5-ae1d-ad7a-0da3-5a7b30526b55.md)** collection that represents all the frames in a document. Read-only.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **Frames**

 _expression_A variable that represents a  ** [Document](8d83487a-2345-a036-a916-971c9db5b7fb.md)** object.


## Remarks
<a name="sectionSection1"> </a>

For information about returning a single member of a collection, see  [Returning an Object from a Collection](28f76384-f495-9640-a7c8-10ada3fac727.md).


## Example
<a name="sectionSection2"> </a>

This example adds a frame around the selection and returns a frame object to the myFrame variable.


```
Set myFrame = ActiveDocument.Frames.Add(Range:=Selection.Range)
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [Document Object](8d83487a-2345-a036-a916-971c9db5b7fb.md)
#### Other resources


 [Document Object Members](fc9ab457-0888-f917-3d52-387168ac23b9.md)
