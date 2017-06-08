---
title: TextRange2.Application Property (Office)
ms.prod: office
api_name:
- Office.TextRange2.Application
ms.assetid: 3883561f-229b-92f9-eaea-83f00ac33f06
ms.date: 06/08/2017
---


# TextRange2.Application Property (Office)

Used without an object qualifier, this property returns an  **Application** object that represents the current instance of the Microsoft Office application. Used with an object qualifier, this property returns an **Application** object that represents the creator of the **TextRange2** object. When used with an OLE Automation object, it returns the object's application. Read-only.


## Syntax

 _expression_. **Application**

 _expression_ An expression that returns a **TextRange2** object.


### Return Value

Object


## Example

This example displays the name of the application that created each linked OLE object on page one of the active Publisher publication.


```
Dim shpOle As Shape 
 
For Each shpOle In ActiveDocument.Pages(1).Shapes 
 If shpOle.Type = pbLinkedOLEObject Then 
 MsgBox shpOle.OLEFormat.Application.Name 
 End If 
Next
```


## See also


#### Concepts


[TextRange2 Object](textrange2-object-office.md)
#### Other resources


[TextRange2 Object Members](textrange2-members-office.md)

