
# PageSetup.LineNumbering Property (Word)

Returns or sets a  **[LineNumbering](a2dd1278-c7dd-af4c-be32-1daded5556d6.md)** object that represents the line numbers for the specified **PageSetup** object.


## Syntax

 _expression_ . **LineNumbering**

 _expression_ An expression that returns a **[PageSetup](1879d601-80ad-4fc0-1a87-92e999b59f88.md)** object.


## Remarks

You must be in print layout view to see line numbering.


## Example

This example enables line numbering for the active document.


```vb
ActiveDocument.PageSetup.LineNumbering.Active = True
```

This example enables line numbering for a document named "MyDocument.doc" The starting number is set to one, every fifth line number is shown, and the numbering is continuous throughout all sections in the document.




```vb
set myDoc = Documents("MyDocument.doc") 
With myDoc.PageSetup.LineNumbering 
 .Active = True 
 .StartingNumber = 1 
 .CountBy = 5 
 .RestartMode = wdRestartContinuous 
End With
```

This example sets the line numbering in the active document equal to the line numbering in MyDocument.doc.




```vb
ActiveDocument.PageSetup.LineNumbering = Documents("MyDocument.doc") _ 
 .PageSetup.LineNumbering
```


## See also


#### Concepts


[PageSetup Object](1879d601-80ad-4fc0-1a87-92e999b59f88.md)
#### Other resources


[PageSetup Object Members](9ff8b896-933b-1a19-19d5-5e5d87aab1b5.md)
