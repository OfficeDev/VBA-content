
# Global.Documents Property (Word)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Returns a  ** [Documents](fc4ac973-19c1-703a-5538-f4426b8b7564.md)** collection that represents all the open documents. Read-only.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **Documents**

 _expression_A variable that represents a  ** [Global](b91e7459-08d5-ea8c-42e0-f7b9bfd1a72c.md)** object.


## Remarks
<a name="sectionSection1"> </a>

For information about returning a single member of a collection, see  [Returning an Object from a Collection](28f76384-f495-9640-a7c8-10ada3fac727.md).


## Example
<a name="sectionSection2"> </a>

This example creates a new document based on the Normal template and then displays the Save As dialog box.


```
Documents.Add.Save
```

This example saves open documents that have changed since they were last saved.




```
Dim docLoop As Document 
 
For Each docLoop In Documents 
 If docLoop.Saved = False Then docLoop.Save 
Next docLoop
```

This example prints each open document after setting the left and right margins to 0.5 inch.




```
Dim docLoop As Document 
 
For Each docLoop In Documents 
 With docLoop 
 .PageSetup.LeftMargin = InchesToPoints(0.5) 
 .PageSetup.RightMargin = InchesToPoints(0.5) 
 .PrintOut 
 End With 
Next docLoop
```

This example opens Doc.doc as a read-only document.




```
Documents.Open FileName:="C:\Files\Doc.doc", ReadOnly:=True
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [Global Object](b91e7459-08d5-ea8c-42e0-f7b9bfd1a72c.md)
#### Other resources


 [Global Object Members](35050f7b-bc46-4795-ec17-f68e263c8af0.md)
