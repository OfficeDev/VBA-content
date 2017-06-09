---
title: Presentation.CustomDocumentProperties Property (PowerPoint)
keywords: vbapp10.chm583021
f1_keywords:
- vbapp10.chm583021
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.CustomDocumentProperties
ms.assetid: 3f972f15-f606-0a11-56b6-1994e617def2
ms.date: 06/08/2017
---


# Presentation.CustomDocumentProperties Property (PowerPoint)

Returns a  **DocumentProperties** collection that represents all the custom document properties for the specified presentation. Read-only.


## Syntax

 _expression_. **CustomDocumentProperties**

 _expression_ A variable that represents a **Presentation** object.


### Return Value

DocumentProperties


## Remarks

Use the  **[BuiltInDocumentProperties](presentation-builtindocumentproperties-property-powerpoint.md)** property to return the collection of built-in document properties.

For information about returning a single member of a collection, see [Returning an Object from a Collection](return-objects-from-collections.md).


## Example

This example adds a static custom property named "Complete" for the active presentation.


```vb
Application.ActivePresentation.CustomDocumentProperties _
    .Add Name:="Complete", LinkToContent:=False, _
    Type:=msoPropertyTypeBoolean, Value:=False
```

This example displays the active presentation if the value of the "Complete" custom property is  **True**.




```vb
With Application.ActivePresentation

    If .CustomDocumentProperties("complete") Then .PrintOut

End With
```


## See also


#### Concepts


[Presentation Object](presentation-object-powerpoint.md)

