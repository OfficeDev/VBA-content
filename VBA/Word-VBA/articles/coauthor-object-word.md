---
title: CoAuthor Object (Word)
keywords: vbawd10.chm1237
f1_keywords:
- vbawd10.chm1237
ms.prod: word
api_name:
- Word.CoAuthor
ms.assetid: d1b58eea-4570-ffd3-4c13-a74a998b079e
ms.date: 06/08/2017
---


# CoAuthor Object (Word)

Represents a single co author in the document. The  **CoAuthor** object is a member of the **[CoAuthors](coauthors-object-word.md)** collection. The **CoAuthors** collection contains all the co authors in the document (authors that are actively editing the document).


 **Important**  Documents can only be co authored on a server that supports the File Synchronization via SOAP over HTTP protocol, such as Microsoft SharePoint Server 2010.


## Remarks

Use  **CoAuthors** ( _Index_ ), where _Index_ is the index number to return a single **CoAuthor** object.


 **Note**  When a new co author begins to edit the document, it can take up to one minute or longer for the co author to appear in the document.


## Example

The following code example returns the name of the first co author in the active document.


```vb
Dim author As CoAuthor 
 
Set author = ActiveDocument.CoAuthoring.Authors(1) 
MsgBox "The name of the first co author in this document is " &; author.Name
```


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


