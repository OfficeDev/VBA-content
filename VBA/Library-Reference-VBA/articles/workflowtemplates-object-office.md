---
title: WorkflowTemplates Object (Office)
keywords: vbaof11.chm283000
f1_keywords:
- vbaof11.chm283000
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.WorkflowTemplates
ms.assetid: 01df4716-4440-7761-8504-22f78e40f8e4
---


# WorkflowTemplates Object (Office)

Represents a collection of  **WorkflowTemplate** objects.


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

