---
title: Document.CheckIn Method (Visio)
keywords: vis_sdr.chm10552030
f1_keywords:
- vis_sdr.chm10552030
ms.prod: visio
api_name:
- Visio.Document.CheckIn
ms.assetid: 9b75d468-24bc-e205-cafa-6e585f469e38
ms.date: 06/08/2017
---


# Document.CheckIn Method (Visio)

Returns a document from a local computer to a Microsoft SharePoint Server computer.


## Syntax

 _expression_ . **CheckIn**( **_SaveChanges_** , **_Comments_** , **_MakePublic_** )

 _expression_ A variable that represents a **Document** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SaveChanges_|Optional| **Boolean**| **True** (non-zero) to save document changes before check-in; **False** (0) to check the document in without saving changes. The default is **True** .|
| _Comments_|Optional| **Variant**|Any comments to be stored with this revision of the document (applies only if  _SaveChanges_ equals **True** ).|
| _MakePublic_|Optional| **Boolean**| **True** to publish the document after check-in. This submits the document for the approval process or, if there is no approval routing for the document, a public version is created that is available to readers of the folder (applies only if _SaveChanges_ equals **True** ); **False** leaves the document available only for private viewing. The default is **False** .|

### Return Value

Nothing


## Remarks

To check in a file, it must be stored in a document library on a computer running Microsoft SharePoint Server or Microsoft SharePoint Foundation.

After the document has been checked in using the  **CheckIn** method, the document is closed. This behavior is different from the user interface; when you check in a document in the user interface, the document is closed and re-opened as read-only.


## Example

This example checks the server to see if the specified document can be checked in. If it can, this example saves and closes the document, and then checks it back into the server.


```vb
Sub CheckDocIn(varDocCheckIn As Variant) 
  
    If Documents.Item(varDocCheckIn).CanCheckin = True Then  
        Documents.Item(varDocCheckIn).CheckIn  
        MsgBox varDocCheckIn &; " has been checked in."  
    Else  
        MsgBox "This file cannot be checked in " _  
            &; "at this time. Please try again later."  
    End If   
  
End Sub
```

To call the preceding  **CheckDocIn** subroutine, use the following subroutine and replace _servername/workspace/drawing.vdx_ with the path to and name of an actual file located on a Microsoft SharePoint Server computer.




```vb
Sub DocIn()  
 
    Call CheckDocIn _  
        (varDocCheckIn:="http://servername/workspace/drawing.vdx ") 
  
End Sub
```


