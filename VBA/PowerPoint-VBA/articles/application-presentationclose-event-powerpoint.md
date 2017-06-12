---
title: Application.PresentationClose Event (PowerPoint)
keywords: vbapp10.chm621004
f1_keywords:
- vbapp10.chm621004
ms.prod: powerpoint
api_name:
- PowerPoint.Application.PresentationClose
ms.assetid: 4057b50a-5f2d-78bf-d55a-d0781da27ea7
ms.date: 06/08/2017
---


# Application.PresentationClose Event (PowerPoint)

Occurs immediately before any open presentation closes, as it is removed from the  **[Presentations](presentations-object-powerpoint.md)** collection.


## Syntax

 _expression_. **PresentationClose**( **_Pres_** )

 _expression_ An expression that returns a **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Pres_|Required|**Presentation**|The presentation that is being closed.|

## Example

This example saves a copy of the active presentation as an HTML file, with the same name and within the same folder.


```vb
Private Sub App_PresentationClose(ByVal Pres As Presentation)
    FindNum = InStr(1, Pres.FullName, ".")
    HTMLName = Mid(Pres.FullName, 1, FindNum - 1) _
        &; ".htm"
    Pres.SaveCopyAs HTMLName, ppSaveAsHTML
End Sub
```


## See also


#### Concepts


[Application Object](application-object-powerpoint.md)

