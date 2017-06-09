---
title: Presentation.CanCheckIn Method (PowerPoint)
keywords: vbapp10.chm583066
f1_keywords:
- vbapp10.chm583066
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.CanCheckIn
ms.assetid: 26d76ca4-4fd3-2037-e193-0d2d39f59361
ms.date: 06/08/2017
---


# Presentation.CanCheckIn Method (PowerPoint)

Returns  **True** if Microsoft PowerPoint can check in a specified presentation to a server.


## Syntax

 _expression_. **CanCheckIn**

 _expression_ A variable that represents a **Presentation** object.


### Return Value

Boolean


## Remarks

To take advantage of the collaboration features built into PowerPoint, you must store presentations on a Microsoft SharePoint Portal Server.


## Example

This example checks the server to see if the specified presentation can be checked in and, if it can be, closes the presentation and checks it back into server.


```vb
Sub CheckInPresentation(strPresentation As String)

    If Presentations(strPresentation).CanCheckIn = True Then
        Presentations(strPresentation).CheckIn
        MsgBox strPresentation &; " has been checked in."
    Else
        MsgBox strPresentation &; " cannot be checked in " &; _
        "at this time.  Please try again later."
    End If

End Sub
```

To call the subroutine above, use the following subroutine and replace the " _http://servername/workspace/report.ppt_ " file name with an actual file located on a server mentioned in the Remarks section above.




```vb
Sub CheckPPTIn()
    Call CheckInPresentation(strPresentation:= _
        "http://servername/workspace/report.ppt ")
End Sub
```


## See also


#### Concepts


[Presentation Object](presentation-object-powerpoint.md)

