---
title: Documents.OpenEx Method (Visio)
keywords: vis_sdr.chm10616405
f1_keywords:
- vis_sdr.chm10616405
ms.prod: visio
api_name:
- Visio.Documents.OpenEx
ms.assetid: 86b26b53-c555-2d36-74d6-0d2a4d81971c
ms.date: 06/08/2017
---


# Documents.OpenEx Method (Visio)

Opens an existing Microsoft Visio file, using extra information passed in as an argument.


## Syntax

 _expression_ . **OpenEx**( **_FileName_** , **_Flags_** )

 _expression_ A variable that represents a **Documents** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FileName_|Required| **String**|The name of the file to open.|
| _Flags_|Required| **Integer**|Flags that indicate how to open the file.|

### Return Value

Document


## Remarks

The  **OpenEx** method is identical to the **Open** method, except that it provides an extra argument in which the caller can specify how the document opens.

The Flags argument should be a combination of zero or more of the following values.



|** Constant**|** Value**|
|:-----|:-----|
| **visOpenCopy**| &;H1|
| **visOpenRO**| &;H2|
| **visOpenDocked**| &;H4|
| **visOpenDontList**| &;H8|
| **visOpenMinimized**| &;H10|
| **visOpenRW**| &;H20|
| **visOpenHidden**| &;H40|
| **visOpenMacrosDisabled**| &;H80|
| **visOpenNoWorkspace**|&;H100|
If  **visOpenDocked** is specified, the file appears in a docked rather than an MDI window, provided that the file is a stencil file and there is an active drawing window in which to put the docked stencil window.

If  **visOpenDontList** is specified, the name of the opened file does not appear in the list of recently opened documents in the **Recent Documents** list on the **Recent** tab (click the **File** tab, and then click **Recent**).

If  **visOpenMinimized** is specified, the file opens minimized?it is not active. This flag is not supported in versions of Visio earlier than 5.0b.

If  **visOpenMacrosDisabled** is specified, the file opens with Visual Basic macros disabled. This flag is not supported in versions earlier than Visio 2002.

If  **visOpenHidden** is specified, the file opens in a hidden window.

If  **visOpenNoWorkspace** is specified, the file opens with no workspace information.

If your Visual Studio solution includes the  **Microsoft.Office.Interop.Visio** reference, this method maps to the following types:


-  **Microsoft.Office.Interop.Visio.IVDocuments.OpenEx(string, short)**
    

## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **OpenEx** method to open a copy of a stencil file in Visio.


```vb
 
Public Sub OpenEx_Example()  
 
    'Use the OpenEx method to open a copy of a stencil.  
    Documents.OpenEx "Basic_U.vss", visOpenDocked + visOpenCopy  
 
End Sub
```


