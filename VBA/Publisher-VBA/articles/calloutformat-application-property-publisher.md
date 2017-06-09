---
title: CalloutFormat.Application Property (Publisher)
keywords: vbapb10.chm2490369
f1_keywords:
- vbapb10.chm2490369
ms.prod: publisher
api_name:
- Publisher.CalloutFormat.Application
ms.assetid: 72ae672f-3234-fbab-274e-6f9d4edcadf1
ms.date: 06/08/2017
---


# CalloutFormat.Application Property (Publisher)

Used without an object qualifier, this property returns an  **[Application](application-object-publisher.md)** object that represents the current instance of Publisher. Used with an object qualifier, this property returns an  **Application** object that represents the creator of the specified object. When used with an OLE Automation object, it returns the object's application.


## Syntax

 _expression_. **Application**

 _expression_A variable that represents a  **CalloutFormat** object.


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


