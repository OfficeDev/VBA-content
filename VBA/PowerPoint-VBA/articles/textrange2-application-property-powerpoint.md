---
title: TextRange2.Application Property (PowerPoint)
ms.assetid: 87be86f1-e5c6-4698-9262-139f7c3e5b44
ms.date: 06/08/2017
ms.prod: powerpoint
---


# TextRange2.Application Property (PowerPoint)

Used without an object qualifier, this property returns an  **Application** object that represents the current instance of the Microsoft Office application. Used with an object qualifier, this property returns an **Application** object that represents the creator of the **TextRange2** object. When used with an OLE Automation object, it returns the object's application. Read-only.


## Syntax

 _expression_. **Application**

 _expression_ An expression that returns a **TextRange2** object.


### Return Value

Object


## Example

This example displays the name of the application that created each linked OLE object on page one of the active Publisher publication.


```vb
Dim shpOle As Shape 
 
For Each shpOle In ActiveDocument.Pages(1).Shapes 
 If shpOle.Type = pbLinkedOLEObject Then 
 MsgBox shpOle.OLEFormat.Application.Name 
 End If 
Next
```


## See also


#### Concepts


[TextRange2 Object (PowerPoint)](textrange2-object-powerpoint.md)


