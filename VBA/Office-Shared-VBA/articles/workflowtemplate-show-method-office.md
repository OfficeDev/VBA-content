---
title: WorkflowTemplate.Show Method (Office)
keywords: vbaof11.chm282006
f1_keywords:
- vbaof11.chm282006
ms.prod: office
api_name:
- Office.WorkflowTemplate.Show
ms.assetid: aa4780b5-f3bd-431f-8cb3-20c6058ebc5a
ms.date: 06/08/2017
---


# WorkflowTemplate.Show Method (Office)

Displays a workflow specific configuration user interface for the specified  **WorkflowTemplate** object.


## Syntax

 _expression_. **Show**

 _expression_ An expression that returns a **WorkflowTemplate** object.


### Return Value

Integer


## Example

The following example displays the name of each workflow template in the current document and then displays workflow specific configuration user interface for a specific template.


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


## See also


#### Concepts


[WorkflowTemplate Object](workflowtemplate-object-office.md)
#### Other resources


[WorkflowTemplate Object Members](workflowtemplate-members-office.md)

