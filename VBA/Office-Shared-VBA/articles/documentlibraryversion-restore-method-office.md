---
title: DocumentLibraryVersion.Restore Method (Office)
keywords: vbaof11.chm277024
f1_keywords:
- vbaof11.chm277024
ms.prod: office
api_name:
- Office.DocumentLibraryVersion.Restore
ms.assetid: 1f6bb17f-a6b7-c52b-7880-9b3f2ed7ff13
ms.date: 06/08/2017
---


# DocumentLibraryVersion.Restore Method (Office)

Restores a previous saved version of a shared document from the  **DocumentLibraryVersions** collection.


## Syntax

 _expression_. **Restore**

 _expression_ A variable that represents a **DocumentLibraryVersion** object.


### Return Value

Object


## Remarks

Use the  **Restore** method to return to an earlier saved version of the active document. The **Restore** method does several things:


1. It changes the open version of the shared document to read-only mode but leaves it open.
    
2. It opens the restored version in read/write mode.
    
3. It saves the restored version to the server as a new document version, making the restored version the latest version.
    


The  **Restore** method raises a run-time error if the active document has changes that have not been saved.


## Example

The following example restores the previous version of the active document.


```
 Dim dlvVersions As Office.DocumentLibraryVersions 
 Set dlvVersions = ActiveDocument.DocumentLibraryVersions 
 dlvVersions(dlvVersions.Count - 1).Restore 
 Set dlvVersions = Nothing 

```


## See also


#### Concepts


[DocumentLibraryVersion Object](documentlibraryversion-object-office.md)
#### Other resources


[DocumentLibraryVersion Object Members](documentlibraryversion-members-office.md)

