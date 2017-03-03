---
title: WorkflowTemplate Object (Office)
keywords: vbaof11.chm282000
f1_keywords:
- vbaof11.chm282000
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.WorkflowTemplate
ms.assetid: 965d0474-dd51-9b0e-b34c-a11f921ff410
---


# WorkflowTemplate Object (Office)

Represents one of the workflows available for the current document.


## Remarks

A  **WorkflowTemplate** object corresponds to one of the options displayed in the **Start New Workflow** dialog box. On a Web page, the workflow templates are displayed as a list of options.


## Example

The following example displays the name of each workflow template in the current document and then displays workflow specific configuration user interface for a specific template. It should be noted that calling the  **GetWorkflowTemplates** method involves a round-trip to the server.


```vb
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


## See also


#### Concepts


[Object Model Reference](../../Office-Shared-VBA/articles/reference-object-library-reference-for-office.md)

