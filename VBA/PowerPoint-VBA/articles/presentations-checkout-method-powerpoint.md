---
title: Presentations.CheckOut Method (PowerPoint)
keywords: vbapp10.chm522007
f1_keywords:
- vbapp10.chm522007
ms.prod: powerpoint
api_name:
- PowerPoint.Presentations.CheckOut
ms.assetid: c6145ab1-f6d5-265a-8244-40b5fa67aedf
ms.date: 06/08/2017
---


# Presentations.CheckOut Method (PowerPoint)

Copies a specified presentation from a server to a local computer for editing. Returns a  **String** that represents the local path and file name of the presentation checked out.


## Syntax

 _expression_. **CheckOut**( **_FileName_** )

 _expression_ A variable that represents a **Presentations** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FileName_|Required|**String**|The server path and name of the presentation.|

### Return Value

Nothing


## Remarks

To take advantage of the collaboration features built into Microsoft PowerPoint, presentations must be stored on a Microsoft Office SharePoint Portal Server.


## Example

This example verifies that a presentation is not checked out by another user and that it can be checked out. If the presentation can be checked out, it copies the presentation to the local computer for editing.


```vb
Sub CheckOutPresentation(strPresentation As String)

    Dim strFileName As String

    With Presentations
        If .CanCheckOut(strPresentation) = True Then
            .CheckOut FileName:=strPresentation
            .Open FileName:=strFileName
        Else
            MsgBox "You are unable to check out this " &; _
                "presentation at this time."
        End If
    End With
End Sub
```

To call the subroutine above, use the following subroutine and replace the  _"http://servername/workspace/report.ppt"_ file name for an actual file located on a server mentioned in the Remarks section above.




```vb
Sub CheckPPTOut()
    Call CheckOutPresentation(strPresentation:= _
        "http://servername/workspace/report.doc")
End Sub
```


## See also


#### Concepts


[Presentations Object](presentations-object-powerpoint.md)

