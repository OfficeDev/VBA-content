---
title: DocumentInspector.Inspect Method (Office)
keywords: vbaof11.chm279003
f1_keywords:
- vbaof11.chm279003
ms.prod: office
api_name:
- Office.DocumentInspector.Inspect
ms.assetid: 5973fa7d-7218-74e3-b67c-c03fbaf4b930
ms.date: 06/08/2017
---


# DocumentInspector.Inspect Method (Office)

Inspects a document for specific information or document properties.


## Syntax

 _expression_. **Inspect**( **_Status_**, **_Results_** )

 _expression_ An expression that returns a **DocumentInspector** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Status_|Required|**MsoDocInspectorStatus**|An enumeration representing that status of the document. Status is an output parameter which means that its value is returned when the method has completed its purpose.|
| _Results_|Required|**String**|Contains a lists the information items or document properties found in the document.|

## Remarks

MsoDocInspectorStatus members


## Example

The following example inspects a document using  **Inspect** method of the **DocumentInspector** object and then displays the status and results of the inspection.


```
Public Sub DI_InspectDocument() 
Dim docStatus As MsoDocInspectorStatus 
Dim result As String 
ActiveDocument.DocumentInspectors(1).Inspect docStatus, results 
 
MsgBox ("The inspection returned the following status " &amp; docStatus &amp; _ 
" with this result " &amp; result) 
End Sub
```


## See also


#### Concepts


[DocumentInspector Object](documentinspector-object-office.md)
#### Other resources


[DocumentInspector Object Members](documentinspector-members-office.md)

