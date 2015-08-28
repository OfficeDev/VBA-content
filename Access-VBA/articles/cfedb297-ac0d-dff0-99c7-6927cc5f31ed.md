
# Name Object (Excel)

 **Last modified:** July 28, 2015

Represents a defined name for a range of cells. Names can be either built-in names â€” such as Database, Print_Area, and Auto_Open â€” or custom names.

## Remarks

 **Application, Workbook, and Worksheet Objects**

The  **Name** object is a member of the ** [Names](ffecf89d-7bae-c470-8e37-608857a9de2a.md)** collection for the ** [Application](19b73597-5cf9-4f56-8227-b5211f657f6f.md)**,  ** [Workbook](8c00aa60-c974-eed3-0812-3c9625eb0d4c.md)**, and  ** [Worksheet](182b705e-854a-81cc-a4b0-59b942de55ae.md)** objects. Use ** [Names](26be56ec-ea12-1600-602a-eb338d4a5a8b.md)**( _index_), where  _index_ is the name index number or defined name, to return a single **Name** object.

The index number indicates the position of the name within the collection. Names are placed in alphabetic order, from a to z, and are not case-sensitive.

 **Range Objects**

Although a  ** [Range](b8207778-0dcc-4570-1234-f130532cc8cd.md)** object can have more than one name, there's no **Names** collection for the **Range** object. Use ** [Name](39d1a326-e123-443c-29c0-453f7b4a8760.md)** with a **Range** object to return the first name from the list of names (sorted alphabetically) assigned to the range. The following example sets the ** [Visible](48860564-6079-932e-2cae-0802235be61e.md)** property for the first name assigned to cells A1:B1 on worksheet one.


## Example

The following example displays the cell reference for the first name in the application collection.


```
MsgBox Names(1).RefersTo
```

The following example deletes the name "mySortRange" from the active workbook.




```
ActiveWorkbook.Names("mySortRange").Delete
```

Use the  **Name** property to return or set the text of the name itself. The following example changes the name of the first **Name** object in the active workbook.




```
Names(1).Name = "stock_values"
```

The following example sets the  **Visible** property for the first name assigned to cells A1:B1 on worksheet one.




```
Worksheets(1).Range("a1:b1").Name.Visible = False
```


## See also


#### Concepts


 [Excel Object Model Reference](11ea8598-8a20-92d5-f98b-0da04263bf2c.md)
#### Other resources


 [Name Object Members](7c35e8e8-4f81-7cec-da3e-faf738903726.md)
