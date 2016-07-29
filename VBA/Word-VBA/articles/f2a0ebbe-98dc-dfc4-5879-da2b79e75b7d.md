
# Document.Container Property (Word)

Returns the object that represents the container application for the specified document. Read-only  **Object** .


## Syntax

 _expression_ . **Container**

 _expression_ A variable that represents a **[Document](8d83487a-2345-a036-a916-971c9db5b7fb.md)** object.


## Remarks

The  **Container** property provides access to the specified document's container application if the document is embedded in another application as an OLE object. This property also provides a pathway into the object model of the container application if a Word document is opened as an ActiveX document — for example, when a Word document is opened in Microsoft Office Binder or Internet Explorer.


## Example

This example displays the name of the container application for the first shape in the active document. For the example to work, this shape must be an OLE object.


```
Msgbox ActiveDocument.Shapes(1).OLEFormat.Object.Container.Name
```


## See also


#### Concepts


[Document Object](8d83487a-2345-a036-a916-971c9db5b7fb.md)
#### Other resources


[Document Object Members](fc9ab457-0888-f917-3d52-387168ac23b9.md)
