---
title: Broadcast.AddMeetingNotes Method (Word)
keywords: vbawd10.chm36438121
f1_keywords:
- vbawd10.chm36438121
ms.prod: word
ms.assetid: e13a52fd-d0a4-bc32-2d0a-f01f9218bfa2
ms.date: 06/08/2017
---


# Broadcast.AddMeetingNotes Method (Word)

Adds shared meeting notes for the specified broadcast that are accessible to attendees who use either Microsoft OneNote 2013 rich client or web app.


## Syntax

 _expression_ . **AddMeetingNotes**_(notesUrl,_ _notesWacUrl)_

 _expression_ A variable that represents a **Broadcast** object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
|||||
| _notesUrl_|Required|STRING|Specifies the URL where the shared meeting notes are stored, for attendees using the Microsoft OneNote 2013 rich client.|
| _notesWacUrl_|Required|STRING|Specifies the URL where the shared meeting notes are stored, for attendees using the Microsoft OneNote 2013 web access client.|

### Return value

 **VOID**


## Remarks

If you fail to pass a string for either of the two parameters, the  **AddMeetingNotes** method returns an Invalid Parameter error. If for any reason the method call fails, Word returns a generic broadcast error.


## See also


#### Other resources


[Broadcast Object](broadcast-object-word.md)


