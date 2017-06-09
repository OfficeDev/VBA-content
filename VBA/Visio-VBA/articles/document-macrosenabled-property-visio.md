---
title: Document.MacrosEnabled Property (Visio)
keywords: vis_sdr.chm10552080
f1_keywords:
- vis_sdr.chm10552080
ms.prod: visio
api_name:
- Visio.Document.MacrosEnabled
ms.assetid: 361b7bad-55f9-2d4b-4de3-8a12da48d59e
ms.date: 06/08/2017
---


# Document.MacrosEnabled Property (Visio)

Specifies whether you can execute macros and process events in a document's Microsoft Visual Basic for Applications (VBA) project. Read-only.


## Syntax

 _expression_ . **MacrosEnabled**

 _expression_ A variable that represents a **Document** object.


### Return Value

Boolean


## Remarks

If your document contains macros that are necessary to your solution's execution, you can use the  **MacrosEnabled** property to verify whether macros are enabled in the document. If they are disabled, you can display a message indicating that your solution may not work as expected because document settings prohibit macros from being executed.

The value of the  **MacrosEnabled** property depends on a combination of the macro setting and the project's signature status (whether it is digitally signed by a trusted source or in a trusted location). The following table describes these combinations.



|**Macro setting**|**Digitally signed **|**In a trusted location**|**MacrosEnabled property**|
|:-----|:-----|:-----|:-----|
| **Disable all macros without notification**|N/A|No|False|
| **Disable all macros without notification**|N/A|Yes|True|
| **Disable all macros with user notification**|N/A|No|False|
| **Disable all macros with user notification**|N/A|Yes|True|
| **Disable all macros except digitally signed macros**|No|No|False|
| **Disable all macros except digitally signed macros**|Yes|N/A|True|
| **Disable all macros except digitally signed macros**|N/A|Yes|True|
| **Enable all macros**|N/A|N/A|True|
By default, macros are disabled in a Visio document not from a trusted publisher, or that is not digitally signed, or that is not in a trusted location.

However, you can change default settings in the  **Macro Settings** category of the Visio **Trust Center** (click the **File** tab, click **Options**, click  **Trust Center**, and then click  **Trust Center Settings**). If  **Disable all macros except digitally signed macros** is selected, macros in Visio documents not in a trusted location are enabled only if the documents are digitally signed. If you select **Disable all macros without notification** or **Disable all macros with notification**, macros in documents not in a trusted location are disabled. If you select  **Enable all macros**, all macros are always enabled, but this option presents a security risk and is not recommended.

Trusted sources are listed in the  **Trusted Publishers** category in the **Trust Center**, and trusted locations are listed in the  **Trusted Locations** category.

To open a document in a disabled state (macros are not enabled), you can use the  **OpenEx** method of the **Document** object. For example:




```
Documents.OpenEx(fileName , visOpenMacrosDisabled)
```


## Example

The following example shows how use to open a document from an add-on and use the  **MacrosEnabled** property to determine whether macros are enabled. If macros are disabled, a message box appears warning the user of limited functionality. Before running this example, supply a valid document file name for the variable _filename_ .


```vb
 
Public Sub MacrosEnabled_Example() 
 
    Dim vsoDocument As Visio.Document 
    Dim blsStatus As Boolean 
 
    Set vsoDocument = Documents.Open("filename ") 
    blsStatus = vsoDocument.MacrosEnabled 
 
    If Not blsStatus Then 
 
         MsgBox "Macro execution has been disabled for this document." &; _  
            "Functionality may be limited." 
 
    End if 
 
End Sub
```


