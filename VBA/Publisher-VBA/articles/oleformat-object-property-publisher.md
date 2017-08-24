---
title: OLEFormat.Object Property (Publisher)
keywords: vbapb10.chm4456451
f1_keywords:
- vbapb10.chm4456451
ms.prod: publisher
api_name:
- Publisher.OLEFormat.Object
ms.assetid: c6bc20e4-4578-7aa1-8cd8-8315b76b28c9
ms.date: 06/08/2017
---


# OLEFormat.Object Property (Publisher)

Returns an  **Object** that represents the specified OLE object's top-level interface. This property allows you to access the properties and methods of an ActiveX control or the application in which an OLE object was created. The OLE object must support OLE Automation for this property to work. Read-only.


## Syntax

 _expression_. **Object**

 _expression_A variable that represents an  **OLEFormat** object.


### Return Value

Object


## Example

This example sets the value of the first shape in the active publication. For the example to work, this first shape must be an ActiveX control (for example, a check box or an option button).


```vb
Dim myObj As Object 
 
With ActiveDocument.Pages(1).Shapes(1).OLEFormat 
 .Activate 
 Set myObj = .Object 
End With 
 
myObj.Value = True
```


