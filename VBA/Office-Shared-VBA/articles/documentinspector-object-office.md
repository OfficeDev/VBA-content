---
title: DocumentInspector Object (Office)
keywords: vbaof11.chm279000
f1_keywords:
- vbaof11.chm279000
ms.prod: office
api_name:
- Office.DocumentInspector
ms.assetid: 75dcf0ca-5afa-996b-e8d2-13d71ac0f6f8
ms.date: 06/08/2017
---


# DocumentInspector Object (Office)

Represents a Document Inspector module in a  **[DocumentInspectors](documentinspectors-object-office.md)** collection.


## Remarks

 **DocumentInspector** object provides access to the **[Inspect](documentinspector-inspect-method-office.md)** and **[Fix](documentinspector-fix-method-office.md)** methods. The **DocumentInspector** object represents custom Document Inspector modules and some "built-in" options. To see the list of built-in options, open the **Document Inspector** dialog box from the **File** menu. The first two options ( **Comments, Revisions, Versions, and Annotations**; **Document Properties and Personal Information** ) are not listed in the **DocumentInspectors** collection; instead their functionality is available through the **RemoveDocumentInformation** method. The remaining options in the dialog box and any installed custom modules developed by third-party sources are available from the **DocumentInspectors** collection by specifying an index value.


## Example

The following example inspects a document using a Document Inspector module and displays the status and results of the inspection.


```
Public Sub DI_InspectDocument() 
Dim docStatus As MsoDocInspectorStatus 
Dim result As String 
ActiveDocument.DocumentInspectors(3).Inspect docStatus, results 
 
MsgBox ("The inspection returned the following status " &amp; docStatus &amp; _ 
" with this result " &amp; result) 
End Sub
```


## Methods



|**Name**|
|:-----|
|[Fix](documentinspector-fix-method-office.md)|
|[Inspect](documentinspector-inspect-method-office.md)|

## Properties



|**Name**|
|:-----|
|[Application](documentinspector-application-property-office.md)|
|[Creator](documentinspector-creator-property-office.md)|
|[Description](documentinspector-description-property-office.md)|
|[Name](documentinspector-name-property-office.md)|
|[Parent](documentinspector-parent-property-office.md)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
