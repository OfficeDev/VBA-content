---
title: Presentations.CanCheckOut Method (PowerPoint)
keywords: vbapp10.chm522008
f1_keywords:
- vbapp10.chm522008
ms.prod: powerpoint
api_name:
- PowerPoint.Presentations.CanCheckOut
ms.assetid: 60393f0c-11e1-169d-2ead-c6556f1d1364
ms.date: 06/08/2017
---


# Presentations.CanCheckOut Method (PowerPoint)

Returns  **True** if Microsoft PowerPoint can check out a specified presentation from a server.


## Syntax

 _expression_. **CanCheckOut**( **_FileName_** )

 _expression_ A variable that represents a **Presentations** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FileName_|Required|**String**|The server path and name of the presentation.|

### Return Value

Boolean


## Remarks

To take advantage of the collaboration features built into PowerPoint, you must store presentations on a Microsoft Office SharePoint Portal Server.


## Example

This example verifies that a presentation is not checked out by another user and that it can be checked out. If the presentation can be checked out, it copies the presentation to the local computer for editing.


```vb
Sub CheckOutPresentation(strPresentation As String)
    If Presentations.CanCheckOut(strPresentation) = True Then
        Presentations.CheckOut FileName:=strPresentation
    Else
        MsgBox "You are unable to check out this " &; _
        "presentation at this time."
    End If
End Sub
```

To call the subroutine above, use the following subroutine and replace the " _http://servername/workspace/report.ppt_ " file name with an actual file located on a server mentioned in the Remarks section above.




```vb
Sub CheckPPTOut()
    Call CheckOutPresentation(strPresentation:= _
        "http://servername/workspace/report.doc ")
End Sub
```


## See also


#### Concepts


[Presentations Object](presentations-object-powerpoint.md)

