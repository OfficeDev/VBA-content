---
title: Presentation.BuiltInDocumentProperties Property (PowerPoint)
keywords: vbapp10.chm583020
f1_keywords:
- vbapp10.chm583020
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.BuiltInDocumentProperties
ms.assetid: d59341c4-70f4-b9be-0db6-3673d588a6bd
ms.date: 06/08/2017
---


# Presentation.BuiltInDocumentProperties Property (PowerPoint)

Returns a  **DocumentProperties** collection that represents all the built-in document properties for the specified presentation. Read-only.


## Syntax

 _expression_. **BuiltInDocumentProperties**

 _expression_ A variable that represents a **Presentation** object.


### Return Value

DocumentProperties


## Remarks

Use the  **[CustomDocumentProperties](presentation-customdocumentproperties-property-powerpoint.md)** property to return the collection of custom document properties.

For information about returning a single member of a collection, see [Returning an Object from a Collection](return-objects-from-collections.md).


## Example

This example displays the names of all the built-in document properties for the active presentation.


```vb
For Each p In Application.ActivePresentation _
        .BuiltInDocumentProperties
    bidpList = bidpList &; p.Name &; Chr$(13)
Next

MsgBox bidpList
```

This example sets the "Category" built-in property for the active presentation if the author of the presentation is Jake Jarmel.




```vb
With Application.ActivePresentation.BuiltInDocumentProperties

    If .Item("author").Value = "Jake Jarmel" Then

        .Item("category").Value = "Creative Writing"

    End If

End With
```


## See also


#### Concepts


[Presentation Object](presentation-object-powerpoint.md)

