---
title: Document.ContentControlOnExit Event (Word)
keywords: vbawd10.chm4001012
f1_keywords:
- vbawd10.chm4001012
ms.prod: word
api_name:
- Word.Document.ContentControlOnExit
ms.assetid: 1c988334-2bb3-2a86-747b-0d1d46894da1
ms.date: 06/08/2017
---


# Document.ContentControlOnExit Event (Word)

Occurs when a user leaves a content control.


## Syntax

Private Sub  _expression_ _**ContentControlOnExit**( **_ContentControl_** , **_Cancel_** )

 _expression_ An expression that returns a **[Document](document-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ContentControl_|Required| **ContentControl**|The content control that the user is leaving.|
| _Cancel_|Required| **Boolean**|Specifies whether to cancel the event.  **True** cancels the event and does not allow the user to leave the control.|

## Remarks


 **Important**  This event fires only for the content control that you exit and not for parent content controls. For example, if you have a text box content control nested inside a group content control, and you move the cursor from inside the text box content control and into another place in the document, this event fires only once for the text box content control; it does not fire for the parent group content control.

For information about using events with the  **Document** object, see[Using Events with the Document Object](http://msdn.microsoft.com/library/2b043342-436a-5421-e8af-3c2c49684960%28Office.15%29.aspx).


## See also


#### Concepts


[Document Object](document-object-word.md)

