---
title: WorkflowTemplate Object (Office)
keywords: vbaof11.chm282000
f1_keywords:
- vbaof11.chm282000
ms.prod: office
api_name:
- Office.WorkflowTemplate
ms.assetid: 965d0474-dd51-9b0e-b34c-a11f921ff410
ms.date: 06/08/2017
---


# WorkflowTemplate Object (Office)

Represents one of the workflows available for the current document.


## Remarks

A  **WorkflowTemplate** object corresponds to one of the options displayed in the **Start New Workflow** dialog box. On a Web page, the workflow templates are displayed as a list of options.


## Example

The following example displays the name of each workflow template in the current document and then displays workflow specific configuration user interface for a specific template. It should be noted that calling the  **GetWorkflowTemplates** method involves a round-trip to the server.


```
Sub DisplayWorkTemplates() 
Dim objWorkflowTemplates As WorkflowTemplates 
Dim objWorkflowTemplate As WorkflowTemplate 
Dim cnt As Integer 
 
Set objWorkflowTemplates = Document.GetWorkflowTemplates() 
 
For cnt = 1 To objWorkflowTemplates.Count 
 Debug.Print objWorkflowTemplate(cnt).Name 
Next 
 
Set objWorkflowTemplate = objWorkflowTemplates(1) 
objWorkflowTemplate.Show 
 
End Sub 

```


## Methods



|**Name**|
|:-----|
|[Show](workflowtemplate-show-method-office.md)|

## Properties



|**Name**|
|:-----|
|[Application](workflowtemplate-application-property-office.md)|
|[Creator](workflowtemplate-creator-property-office.md)|
|[Description](workflowtemplate-description-property-office.md)|
|[DocumentLibraryName](workflowtemplate-documentlibraryname-property-office.md)|
|[DocumentLibraryURL](workflowtemplate-documentlibraryurl-property-office.md)|
|[Id](workflowtemplate-id-property-office.md)|
|[Name](workflowtemplate-name-property-office.md)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
