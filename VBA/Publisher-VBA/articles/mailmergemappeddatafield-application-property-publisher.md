---
title: MailMergeMappedDataField.Application Property (Publisher)
keywords: vbapb10.chm6553601
f1_keywords:
- vbapb10.chm6553601
ms.prod: publisher
api_name:
- Publisher.MailMergeMappedDataField.Application
ms.assetid: 6ac91046-d66d-8e76-c322-a13ea3f1d6c2
ms.date: 06/08/2017
---


# MailMergeMappedDataField.Application Property (Publisher)

Used without an object qualifier, this property returns an  **[Application](application-object-publisher.md)** object that represents the current instance of Publisher. Used with an object qualifier, this property returns an  **Application** object that represents the creator of the specified object. When used with an OLE Automation object, it returns the object's application.


## Syntax

 _expression_. **Application**

 _expression_A variable that represents a  **MailMergeMappedDataField** object.


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


