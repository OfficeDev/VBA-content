---
title: Document.CanCheckIn Method (Visio)
keywords: vis_sdr.chm10552020
f1_keywords:
- vis_sdr.chm10552020
ms.prod: visio
api_name:
- Visio.Document.CanCheckIn
ms.assetid: 99922339-631b-f60e-1d07-3ae9df134cf7
ms.date: 06/08/2017
---


# Document.CanCheckIn Method (Visio)

Specifies whether a document can be checked into a Microsoft SharePoint Server computer.


## Syntax

 _expression_ . **CanCheckIn**

 _expression_ A variable that represents a **Document** object.


### Return Value

Boolean


## Remarks

You cannot check in a file that is not in a checked-out state or stored in a document library on a computer running Microsoft SharePoint Server or Microsoft SharePoint Foundation.


## Example

This example checks the server to see if the specified document can be checked in, and if it can, checks it back into the server.


```vb
 
Sub CheckDocIn (varDocCheckIn As Variant) 
 
 If Documents.Item(varDocCheckIn).CanCheckIn = True Then 
 Documents.Item(varDocCheckIn).CheckIn 
 MsgBox varDocCheckIn &; " has been checked in." 
 Else 
 MsgBox "This file cannot be checked in " &; _ 
 "at this time. Please try again later." 
 End If 
End Sub
```

To call the preceding  **CheckDocIn** subroutine, use the following subroutine and replace _servername/workspace/drawing.vdx_ with the path to and name of an actual file located on a Microsoft SharePoint Server computer.




```vb
 
Sub DocIn() 
 Call CheckDocIn _ 
 (varDocCheckIn:="http://servername/workspace/drawing.vdx ") 
End Sub
```


