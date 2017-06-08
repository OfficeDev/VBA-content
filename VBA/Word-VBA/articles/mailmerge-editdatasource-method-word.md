---
title: MailMerge.EditDataSource Method (Word)
keywords: vbawd10.chm153092203
f1_keywords:
- vbawd10.chm153092203
ms.prod: word
api_name:
- Word.MailMerge.EditDataSource
ms.assetid: 2d1c681e-b8de-4692-288c-7a5b9f501288
ms.date: 06/08/2017
---


# MailMerge.EditDataSource Method (Word)

Opens or switches to the mail merge data source.


## Syntax

 _expression_ . **EditDataSource**

 _expression_ Required. A variable that represents a **[MailMerge](mailmerge-object-word.md)** object.


## Remarks

If the data source is a Microsoft Word document, this method opens the data source (or activates the data source if it is already open).

If Word is accessing the data through dynamic data exchange (DDE)—using an application such as Microsoft Office Excel or Microsoft Office Access—this method displays the data source in that application.


 **Security Note**  



If Word is accessing the data through open database connectivity (ODBC), this method displays the data in a Word document. Note that, if Microsoft Query is installed, a message appears providing the option to display Microsoft Query instead of converting data.


## Example

This example opens or activates the data source attached to the document named "Sales.doc."


```
Documents("Sales.doc").MailMerge.EditDataSource
```

This example opens or activates the attached data source if the data source is a Word document.




```vb
Dim dsMain As MailMergeDataSource 
 
Set dsMain = ActiveDocument.MailMerge.DataSource 
If dsMain.Type = wdMergeInfoFromWord Then 
 ActiveDocument.MailMerge.EditDataSource 
End If
```


## See also


#### Concepts


[MailMerge Object](mailmerge-object-word.md)

