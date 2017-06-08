---
title: Document.CheckInWithVersion Method (Word)
keywords: vbawd10.chm158007797
f1_keywords:
- vbawd10.chm158007797
ms.prod: word
api_name:
- Word.Document.CheckInWithVersion
ms.assetid: fc041188-438e-6fab-d096-7883074a6879
ms.date: 06/08/2017
---


# Document.CheckInWithVersion Method (Word)

Saves a document to a server from a local computer, and sets the local document to read-only so that it cannot be edited locally.


## Syntax

 _expression_ . **CheckInWithVersion**( **_SaveChanges_** , **_Comments_** , **_MakePublic_** , **_VersionType_** )

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SaveChanges_|Optional| **Boolean**| **True** to save the document to the server location. The default is **True** .|
| _Comments_|Optional| **Variant**|Comments for the revision of the document being checked in (applies only if SaveChanges is set to  **True** ).|
| _MakePublic_|Optional| **Boolean**| **True** to allow the user to publish the document after it is checked in.|
| _VersionType_|Optional| **Variant**|Specifies versioning information for the document. |

## Remarks

Setting the MakePublic parameter to  **True** submits the document for the approval process, which can eventually result in a version of the document being published to users with read-only rights to the document (applies only if SaveChanges is set to **True** ).

To take advantage of the collaboration features built into Microsoft Word, documents must be stored on a Microsoft SharePoint Server.


## Example

The following example uses the  **[CanCheckin](document-cancheckin-method-word.md)** method to determine whether the document has been stored on a Microsoft SharePoint Server. If the document has been stored on a server, the example calls the **CheckInWithVersion** method to check in the document along with the specified comments and version number, save changes to the server location, and submit the document for the approval process.

This example is for a document-level customization.




```vb
Private Sub DocumentCheckIn() 
 If ActiveDocument.CanCheckin Then 
 ActiveDocument.CheckInWithVersion _ 
 True, _ 
 "My updates.", _ 
 True, _ 
 WdCheckInVersionType.wdCheckInMinorVersion 
 Else 
 MessageBox.Show ("This document cannot be checked in") 
 End If 
End Sub
```


## See also


#### Concepts


[Document Object](document-object-word.md)

