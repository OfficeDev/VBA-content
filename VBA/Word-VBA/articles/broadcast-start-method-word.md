---
title: Broadcast.Start Method (Word)
keywords: vbawd10.chm36438117
f1_keywords:
- vbawd10.chm36438117
ms.prod: word
ms.assetid: 0a49bf9f-4975-3309-0c23-c758b1aab566
ms.date: 06/08/2017
---


# Broadcast.Start Method (Word)

Initiates the specified broadcast session.


## Syntax

 _expression_ . **Start**_(serverUrl)_

 _expression_ A variable that represents a **Broadcast** object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
|||||
| _serverUrl_|Required|STRING|The URL of the broadcast server.|

### Return value

 **VOID**


## Remarks

Calling the  **Start** method sets up the server, authenticates the user, and uploads the presentation.

If the value passed for  _serverUrl_ has invalid formatting, **Start** returns an ?Invalid Parameter? error. Additionally, the method returns an error if the document is DRM protected, is already being broadcast, or has conflicting edits (is in merge mode).


## See also


#### Other resources


[Broadcast Object](broadcast-object-word.md)


