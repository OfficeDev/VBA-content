---
title: Document.AutoRecover Property (Visio)
keywords: vis_sdr.chm10514710
f1_keywords:
- vis_sdr.chm10514710
ms.prod: visio
api_name:
- Visio.Document.AutoRecover
ms.assetid: 23b09910-35a8-93bc-71f0-4618b1c48523
ms.date: 06/08/2017
---


# Document.AutoRecover Property (Visio)

Determines whether an open document that has unsaved changes is copied when automatic recovery is enabled. Read/write.


## Syntax

 _expression_ . **AutoRecover**

 _expression_ A variable that represents a **Document** object.


### Return Value

Boolean


## Remarks

If automatic recovery is enabled (if the  **Application.AutoRecoverInterval** property is greater than 0), all documents that are open and have unsaved changes are copied into temporary files. If you do not want a document to be recovered, set its **AutoRecover** property to **False** . The **AutoRecover** property is not saved with a document and must be set each time the document opens.

When Microsoft Visio is launched after an abnormal termination and determines that automatic recovery was enabled, it attempts to open all files that were open at termination.




- If there is a recovery file that is more recent than the last saved copy of the file, Visio opens the recovered file and displays the name "<file name> (Recovered)" in the document's title bar.
    
- If there is no recovery file, Visio opens the last saved copy of the document.
    


You must still save changes to recovered documents before Visio closes. If you do not save recovered documents, changes are discarded, as in any unsaved document.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **AutoRecover** property to disable automatic recovery for a particular document.


```vb
 
Private Sub Document_DocumentOpened(ByValdoc As IVDocument) 
  
    'Do not recover this document 
    ThisDocument.AutoRecover = False 
 
End Sub
```


