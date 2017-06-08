---
title: Document.SendForReview Method (Word)
keywords: vbawd10.chm158007649
f1_keywords:
- vbawd10.chm158007649
ms.prod: word
api_name:
- Word.Document.SendForReview
ms.assetid: 2f2cdd5c-eeca-d03f-bd58-b5586f8f461f
ms.date: 06/08/2017
---


# Document.SendForReview Method (Word)

Sends a document in an e-mail message for review by the specified recipients.


## Syntax

 _expression_ . **SendForReview**( **_Recipients_** , **_Subject_** , **_ShowMessage_** , **_IncludeAttachment_** )

 _expression_ Required. A variable that represents a **[Document](document-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Recipients_|Optional| **Variant**|A string that lists the people to whom to send the message. These can be unresolved names and aliases in an e-mail phone book or full e-mail addresses. Separate multiple recipients with a semicolon (;). If left blank and ShowMessage is  **False** , you will receive an error message and the message will not be sent.|
| _Subject_|Optional| **Variant**|A string for the subject of the message. If left blank, the subject will be: Please review "file name".|
| _ShowMessage_|Optional| **Variant**|A  **Boolean** value that indicates whether the message should be displayed when the method is executed. The default value is **True** . If set to **False** , the message is automatically sent to the recipients without first showing the message to the sender.|
| _IncludeAttachment_|Optional| **Variant**|A  **Boolean** value that indicates whether the message should include an attachment or a link to a server location. The default value is **True** . If set to **False** , the document must be stored at a shared location.|

## Remarks

The  **SendForReview** method starts a collaborative review cycle. Use the **EndReview** method to end a review cycle.


## Example

This example automatically sends the current document as an attachment in an e-mail message to the specified recipients.


```vb
Sub WebReview() 
 ActiveDocument.SendForReview _ 
 Recipients:="someone@example.com; amy jones", _ 
 Subject:="Please review this document.", _ 
 ShowMessage:=False, _ 
 IncludeAttachment:=True 
End Sub
```


## See also


#### Concepts


[Document Object](document-object-word.md)

