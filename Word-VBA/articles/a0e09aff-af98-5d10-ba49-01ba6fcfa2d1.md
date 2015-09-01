
# Document.Tables Property (Word)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Returns a  ** [Table](996b58dd-ebc6-ee30-5bfe-c5e51a0f71d6.md)** collection that represents all the tables in the specified document. Read-only.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **Tables**

 _expression_A variable that represents a  ** [Document](8d83487a-2345-a036-a916-971c9db5b7fb.md)** object.


## Remarks
<a name="sectionSection1"> </a>

For information about returning a single member of a collection, see  [Returning an Object from a Collection](28f76384-f495-9640-a7c8-10ada3fac727.md).


## Example
<a name="sectionSection2"> </a>

This example creates a 5x5 table in the active document and then applies a predefined format to it.


```
Selection.Collapse Direction:=wdCollapseStart 
Set myTable = ActiveDocument.Tables.Add(Range:=Selection.Range, _ 
NumRows:=5, NumColumns:=5) 
myTable.AutoFormat Format:=wdTableFormatClassic2
```

This example inserts numbers and text into the first column of the first table in the active document.




```
num = 90 
For Each acell In ActiveDocument.Tables(1).Columns(1).Cells 
 acell.Range.Text = num &amp; " Sales" 
 num = num + 1 
Next acell
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [Document Object](8d83487a-2345-a036-a916-971c9db5b7fb.md)
#### Other resources


 [Document Object Members](fc9ab457-0888-f917-3d52-387168ac23b9.md)
