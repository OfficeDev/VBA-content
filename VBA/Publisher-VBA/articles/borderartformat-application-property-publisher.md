---
title: BorderArtFormat.Application Property (Publisher)
keywords: vbapb10.chm7602177
f1_keywords:
- vbapb10.chm7602177
ms.prod: publisher
api_name:
- Publisher.BorderArtFormat.Application
ms.assetid: 329de5c7-c469-b7c0-1398-cecbbc6107b2
ms.date: 06/08/2017
---


# BorderArtFormat.Application Property (Publisher)

Used without an object qualifier, this property returns an  **[Application](application-object-publisher.md)** object that represents the current instance of Publisher. Used with an object qualifier, this property returns an  **Application** object that represents the creator of the specified object. When used with an OLE Automation object, it returns the object's application.


## Syntax

 _expression_. **Application**

 _expression_A variable that represents a  **BorderArtFormat** object.


## Example

This example displays the version and build information for Publisher.


```vb
With Application 
 MsgBox "Current Publisher: version " _ 
 &; .Version &; " build " &; .Build 
End With
```

This example displays the name of the application that created each linked OLE object on page one of the active publication.




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


 [BorderArtFormat Object](borderartformat-object-publisher.md)

