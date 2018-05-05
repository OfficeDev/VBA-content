---
title: Document.SetLetterContent Method (Word)
keywords: vbawd10.chm158007421
f1_keywords:
- vbawd10.chm158007421
ms.prod: word
api_name:
- Word.Document.SetLetterContent
ms.assetid: 8c9b2f6e-34a7-41a3-761d-c1a5da141aba
ms.date: 06/08/2017
---


# Document.SetLetterContent Method (Word)

Inserts the contents of the specified  **LetterContent** object into a document.


## Syntax

 _expression_ . **SetLetterContent**( **_LetterContent_** )

 _expression_ Required. A variable that represents a **[Document](document-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _LetterContent_|Required| **[LetterContent](lettercontent-object-word.md)**|The **[Document](document-object-word.md)** object that includes the various elements of the letter.|

## Remarks

This method is similar to the  **RunLetterWizard** method except that it doesn't display the Letter Wizard dialog box. The method adds, deletes, or restyles letter elements in the specified document based on the contents of the **LetterContent** object.


## Example

This example retrieves the Letter Wizard elements from the active document, changes the attention line text, and then uses the  **SetLetterContent** method to update the active document to reflect the changes.


```vb
Set myLetterContent = ActiveDocument.GetLetterContent 
myLetterContent.AttentionLine = "Greetings" 
ActiveDocument.SetLetterContent LetterContent:=myLetterContent
```


## See also


#### Concepts


[Document Object](document-object-word.md)

