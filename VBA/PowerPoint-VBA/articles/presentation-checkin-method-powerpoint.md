---
title: Presentation.CheckIn Method (PowerPoint)
keywords: vbapp10.chm583065
f1_keywords:
- vbapp10.chm583065
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.CheckIn
ms.assetid: 63621199-7cda-c464-527f-f55130753f08
ms.date: 06/08/2017
---


# Presentation.CheckIn Method (PowerPoint)

Returns a presentation from a local computer to a server, and sets the local file to read-only so that it cannot be edited locally.


## Syntax

 _expression_. **CheckIn**( **_SaveChanges_**, **_Comments_**, **_MakePublic_** )

 _expression_ A variable that represents a **Presentation** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SaveChanges_|Optional|**Boolean**|**True** saves the presentation to the server location. The default value is **False**.|
| _Comments_|Optional|**Variant**|Comments for the revision of the presentation being checked in (only applies if SaveChanges equals  **True** ).|
| _MakePublic_|Optional|**Variant**|**True** allows the user to perform a publish on the presentation after being checked in. This submits the document for the approval process, which can eventually result in a version of the presentation being published to users with read-only rights to the presentation (only applies if SaveChanges equals **True** ).|

## Remarks

To take advantage of the collaboration features built into Microsoft PowerPoint, presentations must be stored on a Microsoft Office SharePoint Portal Server.


## Example

This example checks the server to see if the specified presentation can be checked in and, if so, closes the presentation and checks it back into server.


```vb
Sub CheckInPresentation(strPresentation As String)

    If Presentations(strPresentation).CanCheckIn = True Then

        Presentations(strPresentation).CheckIn

        MsgBox strPresentation &; " has been checked in."

    Else

        MsgBox strPresentation &; " cannot be checked in at this time.  Please try again later."

    End If

End Sub
```

To call the subroutine above, use the following subroutine and replace the " _http://servername/workspace/report.ppt_ " file name with an actual file located on a server mentioned in the Remarks section above.




```vb
Sub CheckInPresentation()

    Call CheckInPresentation(strPresentation:= "http://servername/workspace/report.ppt ")

End Sub
```


## See also


#### Concepts


[Presentation Object](presentation-object-powerpoint.md)

