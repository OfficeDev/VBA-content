---
title: Application.PresentationSync Event (PowerPoint)
keywords: vbapp10.chm621022
f1_keywords:
- vbapp10.chm621022
ms.prod: powerpoint
api_name:
- PowerPoint.Application.PresentationSync
ms.assetid: 391b486e-7e92-bc90-224a-77c499cdf774
ms.date: 06/08/2017
---


# Application.PresentationSync Event (PowerPoint)

Occurs when the local copy of a presentation that is part of a Document Workspace is synchronized with the copy on the server. Provides important status information regarding the success or failure of the synchronization of the presentation.


## Syntax

 _expression_. **PresentationSync**( **_Pres_**, **_SyncEventType_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Pres_|Required|**Presentation**|The presentation that is being synchronized.|
| _SyncEventType_|Required|**MsoSyncEventType**|The status of the synchronization.|

## Remarks

The  _SyncEventType_ parameter value can be one of these **MsoSyncEventType** constants.


||
|:-----|
|**msoSyncEventDownloadInitiated**|
|**msoSyncEventDownloadSucceeded**|
|**msoSyncEventDownloadFailed**|
|**msoSyncEventUploadInitiated**|
|**msoSyncEventUploadSucceeded**|
|**msoSyncEventUploadFailed**|
|**msoSyncEventDownloadNoChange**|
|**msoSyncEventOffline**|

## Example

The following example displays a message if the synchronization of a presentation in a Document Workspace fails.


```vb
Private Sub app_PresentationSync(ByVal Pres As Presentation, _
        ByVal SyncEventType As Office.MsoSyncEventType)

    If SyncEventType = msoSyncEventDownloadFailed Or _
            SyncEventType = msoSyncEventUploadFailed Then

        MsgBox "Synchronization failed. " &; _
            "Please contact your administrator, " &; vbCrLf &; _
            "or try again later."

    End If

End Sub
```


## See also


#### Concepts


[Application Object](application-object-powerpoint.md)

