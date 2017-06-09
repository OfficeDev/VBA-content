---
title: Documents.CheckOut Method (Visio)
keywords: vis_sdr.chm10652035
f1_keywords:
- vis_sdr.chm10652035
ms.prod: visio
api_name:
- Visio.Documents.CheckOut
ms.assetid: eda3b173-0874-47b6-e18d-a0036e6a31e5
ms.date: 06/08/2017
---


# Documents.CheckOut Method (Visio)

Marks a specified document as checked out and assigns edit privileges to the current user.


## Syntax

 _expression_ . **CheckOut**( **_FileName_** )

 _expression_ A variable that represents a **Documents** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FileName_|Required| **String**|The server path and name of the file.|

### Return Value

Nothing


## Remarks

To check out a file, it must be stored in a document library on a computer running Microsoft SharePoint Server or Microsoft SharePoint Foundation.

Unlike the behavior in the user interface, the  **CheckOut** method does not open the document. Use the **Open** method to open the document in the drawing window after checking it out.


## Example

This example verifies that a document can be checked out. If it can, this example checks it out; otherwise, it returns a message that the document cannot be checked out.


```vb
 
Sub CheckDocOut(strDocCheckOut As String)  
    If Documents.CanCheckOut(strDocCheckOut) = True Then  
        Documents.CheckOut strDocCheckOut  
    Else  
        MsgBox "You are unable to check out this document " _  
            &; "at this time."  
    End If    
End Sub
```

To call the preceding  **CheckDocOut** subroutine, use the following subroutine and replace _servername/workspace/drawing.vdx_ with the path to and name of an actual file located on a Microsoft SharePoint Server computer.




```vb
 
Sub DocOut()  
    Call CheckDocOut _  
        (strDocCheckOut:="http://servername/workspace/drawing.vdx ")  
End Sub
```


