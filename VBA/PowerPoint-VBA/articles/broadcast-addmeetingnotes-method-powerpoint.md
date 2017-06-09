---
title: Broadcast.AddMeetingNotes Method (PowerPoint)
keywords: vbapp10.chm732009
f1_keywords:
- vbapp10.chm732009
ms.assetid: c667cf7c-b4a2-19fc-ad1f-ed8a09c5f769
ms.date: 06/08/2017
ms.prod: powerpoint
---


# Broadcast.AddMeetingNotes Method (PowerPoint)

Adds shared meeting notes for the specified broadcast that are accessible to attendees who use either Microsoft OneNote 2013 rich client or web app.


## Syntax

 _expression_. **AddMeetingNotes**_(notesUrl,_ _notesWacUrl)_

 _expression_ A variable that represents a **Broadcast** object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
|||||
| _notesUrl_|Required|**String**|Specifies the URL where the shared meeting notes are stored, for attendees using the Microsoft OneNote 2013 rich client.|
| _notesWacUrl_|Required|**String**|Specifies the URL where the shared meeting notes are stored, for attendees using the Microsoft OneNote 2013 web access client.|
| _notesUrl_|Required|STRING||
| _notesWacUrl_|Required|STRING||

## Return value

 **VOID**


## Remarks

If you fail to pass a string for either of the two parameters, the  **AddMeetingNotes** method returns an Invalid Parameter error. If for any reason the method call fails, PowerPoint returns a generic broadcast error.


