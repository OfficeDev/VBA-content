---
title: Documents.CanCheckOut Method (Visio)
keywords: vis_sdr.chm10652025
f1_keywords:
- vis_sdr.chm10652025
ms.prod: visio
api_name:
- Visio.Documents.CanCheckOut
ms.assetid: b16fec91-fe8d-3967-7bb2-67ddde44681c
ms.date: 06/08/2017
---


# Documents.CanCheckOut Method (Visio)

Specifies whether a document can be checked out from a Microsoft SharePoint Server computer.


## Syntax

 _expression_ . **CanCheckOut**( **_FileName_** )

 _expression_ A variable that represents a **Documents** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FileName_|Required| **String**|The server path and name of the file.|

### Return Value

Boolean


## Remarks

You cannot check out a file that is already checked out or that is not stored in a document library on a computer running Microsoft SharePoint Server or Microsoft SharePoint Foundation.


## Example

This example verifies that a document is not checked out by another user and that it can be checked out.


```vb
 
Sub CheckDocOut(strDocCheckOut As String) 
 If Documents.CanCheckOut(strDocCheckOut) = True Then 
 Documents.CheckOut strDocCheckOut 
 Else 
 MsgBox "You are unable to check out this " _ 
 &; "document at this time." 
 End If 
End Sub
```

To call the preceding  **CheckDocOut** subroutine, use the following subroutine and replace _servername/workspace/drawing.vdx_ with the path to and name of an actual file located on a Microsoft SharePoint Server computer.




```vb
 
Sub DocOut() 
 Call CheckDocOut _ 
 (strDocCheckOut:="http://servername/workspace/drawing.vdx ") 
End Sub
```


