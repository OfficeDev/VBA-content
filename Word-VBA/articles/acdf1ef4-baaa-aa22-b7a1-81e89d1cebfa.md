
# PageSetup.LineNumbering Property (Word)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Returns or sets a  ** [LineNumbering](a2dd1278-c7dd-af4c-be32-1daded5556d6.md)**object that represents the line numbers for the specified  **PageSetup**object.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **LineNumbering**

 _expression_An expression that returns a  ** [PageSetup](1879d601-80ad-4fc0-1a87-92e999b59f88.md)** object.


## Remarks
<a name="sectionSection1"> </a>

You must be in print layout view to see line numbering.


## Example
<a name="sectionSection2"> </a>

This example enables line numbering for the active document.


```
ActiveDocument.PageSetup.LineNumbering.Active = True
```

This example enables line numbering for a document named "MyDocument.doc" The starting number is set to one, every fifth line number is shown, and the numbering is continuous throughout all sections in the document.




```
set myDoc = Documents("MyDocument.doc") 
With myDoc.PageSetup.LineNumbering 
 .Active = True 
 .StartingNumber = 1 
 .CountBy = 5 
 .RestartMode = wdRestartContinuous 
End With
```

This example sets the line numbering in the active document equal to the line numbering in MyDocument.doc.




```
ActiveDocument.PageSetup.LineNumbering = Documents("MyDocument.doc") _ 
 .PageSetup.LineNumbering
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [PageSetup Object](1879d601-80ad-4fc0-1a87-92e999b59f88.md)
#### Other resources


 [PageSetup Object Members](9ff8b896-933b-1a19-19d5-5e5d87aab1b5.md)
