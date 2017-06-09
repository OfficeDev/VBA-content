---
title: Document.ContentControlOnEnter Event (Word)
keywords: vbawd10.chm4001013
f1_keywords:
- vbawd10.chm4001013
ms.prod: word
api_name:
- Word.Document.ContentControlOnEnter
ms.assetid: 593eca61-886c-85e9-fde2-1dc20c80740b
ms.date: 06/08/2017
---


# Document.ContentControlOnEnter Event (Word)

Occurs when a user enters a content control.


## Syntax

Private Sub  _expression_ _**ContentControlOnEnter**( **_ContentControl_** , )

 _expression_ An expression that returns a **[Document](document-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ContentControl_|Required| **ContentControl**|The content control that the user is entering.|

## Remarks


 **Important**  This event fires only for the content control that you enter and not for parent content controls. For example, if you have a text box content control nested inside a group content control, and you place the cursor inside the text box content control, this event fires only once for the text box content control; it does not fire for the parent group content control.

For information about using events with the  **Document** object, see[Using Events with the Document Object](http://msdn.microsoft.com/library/2b043342-436a-5421-e8af-3c2c49684960%28Office.15%29.aspx).


## See also


#### Concepts


[Document Object](document-object-word.md)

