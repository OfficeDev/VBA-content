---
title: Attachments.Add Method (Publisher)
keywords: vbapb10.chm569349
f1_keywords:
- vbapb10.chm569349
ms.prod: publisher
api_name:
- Publisher.Attachments.Add
ms.assetid: dbf2eb67-5e28-a7e6-226f-feac9045186b
ms.date: 06/08/2017
---


# Attachments.Add Method (Publisher)

Adds an  **Attachment** object to the **Attachments** collection of a Microsoft Publisher publication.


## Syntax

 _expression_. **Add**( **_Filename_**)

 _expression_A variable that represents an  **Attachments** colleciton.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Filename|Required| **String**|File name of the attachment.|

### Return Value

Attachment


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to add an attachment to the message in an e-mail merge. The code adds an attachment to an e-mail message and then prints the number of current attachments to the message in the  **Immediate** window.

The attachment in this example is an image file at the root of the C drive. Before running the code, replace " _C:\image.jpg_" with the path to and name of the file on your computer that you want to add as an e-mail attachment.

Before you can create an e-mail merge, you must use the  **[OpenDataSource](mailmerge-opendatasource-method-publisher.md)** method of the **[MailMerge](mailmerge-object-publisher.md)** object to connect the active document to a data source. To run the merge, use the **[Execute](findreplace-execute-method-publisher.md)** method of the **MailMerge** object. For an example of how to connect to a data source and create an e-mail merge, see the **[EmailMergeEnvelope](emailmergeenvelope-object-publisher.md)** object topic.




```vb
Public Sub Add_Example() 
 
 Dim pubAttachment As Publisher.Attachment 
 
 Set pubAttachment = ThisDocument.MailMerge.EmailMergeEnvelope.Attachemts.Add("C:\image.jpg") 
 Debug.Print ThisDocument.MailMerge.EmailMergeEnvelope.Attachemts.Count 
 
End Sub
```


## See also


#### Concepts


 [Attachments Collection](attachments-object-publisher.md)

