
# OLEFormat.DisplayAsIcon Property (Word)

 **Last modified:** July 28, 2015

 **True** if the specified object is displayed as an icon. Read/write **Boolean**.

## Syntax

 _expression_. **DisplayAsIcon**

 _expression_A variable that represents a  ** [OLEFormat](d4c7aa65-5d3a-0b79-914b-6f908b506f63.md)** object.


## Example

This example displays a message box containing the name of each floating shape that's displayed as an icon on the active document.


```
Dim shapeLoop As Shape 
 
For Each shapeLoop In ActiveDocument.Shapes 
 If shapeLoop.OLEFormat.DisplayAsIcon Then 
 MsgBox shapeLoop.Name &amp; " is displayed as an icon." 
 End If 
Next shapeLoop
```

This example inserts a Microsoft Excel worksheet as a linked OLE object on the active document and then changes the display of the object to an icon.




```
Dim objNew As Object 
 
Set objNew = ActiveDocument.Shapes.AddOLEObject _ 
 (FileName:="C:\Program Files\Microsoft Office" _ 
 &amp; "\Office\Samples\samples.xls", LinkToFile:=True) 
 
objNew.OLEFormat.DisplayAsIcon = True
```


## See also


#### Concepts


 [OLEFormat Object](d4c7aa65-5d3a-0b79-914b-6f908b506f63.md)
#### Other resources


 [OLEFormat Object Members](62aae4c1-c2c6-fbf7-193d-c078ea88a527.md)
