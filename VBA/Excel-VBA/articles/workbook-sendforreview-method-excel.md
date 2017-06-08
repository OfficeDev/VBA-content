---
title: Workbook.SendForReview Method (Excel)
keywords: vbaxl10.chm199206
f1_keywords:
- vbaxl10.chm199206
ms.prod: excel
api_name:
- Excel.Workbook.SendForReview
ms.assetid: 3834f5b3-6d24-1bb9-27b5-052aa2e725e3
ms.date: 06/08/2017
---


# Workbook.SendForReview Method (Excel)

Sends a workbook in an e-mail message for review to the specified recipients.


## Syntax

 _expression_ . **SendForReview**( **_Recipients_** , **_Subject_** , **_ShowMessage_** , **_IncludeAttachment_** )

 _expression_ A variable that represents a **Workbook** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Recipients_|Optional| **Variant**|A string that lists the people to whom to send the message. These can be unresolved names and aliases in an e-mail phone book or full e-mail addresses. Separate multiple recipients with a semicolon (;). If left blank and  _ShowMessage_ is **False** , you will receive an error message, and the message will not be sent.|
| _Subject_|Optional| **Variant**|A string for the subject of the message. If left blank, the subject will be: Please review "filename".|
| _ShowMessage_|Optional| **Variant**|A  **Boolean** value that indicates whether the message should be displayed when the method is executed. The default value is **True** . If set to **False** , the message is automatically sent to the recipients without first showing the message to the sender.|
| _IncludeAttachment_|Optional| **Variant**|A  **Boolean** value that indicates whether the message should include an attachment or a link to a server location. The default value is **True** . If set to **False** , the document must be stored at a shared location.|

## Remarks

The  **SendForReview** method starts a collaborative review cycle. Use the **[EndReview](workbook-endreview-method-excel.md)** method to end a review cycle.


## Example

This example automatically sends the active workbook as an attachment in an e-mail message to the specified recipients.


```vb
Sub WebReview() 
 
 ActiveWorkbook.SendForReview _ 
 Recipients:="someone@example.com; amy jones; lewjudy", _ 
 Subject:="Please review this document.", _ 
 ShowMessage:=False, _ 
 IncludeAttachment:=True 
 
End Sub
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

