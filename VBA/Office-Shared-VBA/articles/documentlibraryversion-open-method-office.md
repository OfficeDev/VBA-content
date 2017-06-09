---
title: DocumentLibraryVersion.Open Method (Office)
keywords: vbaof11.chm277023
f1_keywords:
- vbaof11.chm277023
ms.prod: office
api_name:
- Office.DocumentLibraryVersion.Open
ms.assetid: aa77a821-5fda-209b-a352-81aa9e4fb0d0
ms.date: 06/08/2017
---


# DocumentLibraryVersion.Open Method (Office)

Opens the specified version of the shared document from the  **DocumentLibraryVersions** collection in read-only mode.


## Syntax

 _expression_. **Open**

 _expression_ Required. A variable that represents a **[DocumentLibraryVersion](documentlibraryversion-object-office.md)** object.


## Example

The following example opens the previous saved version of the active document in read-only mode.


```
 Dim dlvVersions As Office.DocumentLibraryVersions 
 Set dlvVersions = ActiveDocument.DocumentLibraryVersions 
 dlvVersions(dlvVersions.Count - 1).Open 
 Set dlvVersions = Nothing 

```


## See also


#### Concepts


[DocumentLibraryVersion Object](documentlibraryversion-object-office.md)
#### Other resources


[DocumentLibraryVersion Object Members](documentlibraryversion-members-office.md)

