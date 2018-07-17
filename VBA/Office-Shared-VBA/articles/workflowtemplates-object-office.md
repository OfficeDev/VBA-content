---
title: WorkflowTemplates Object (Office)
keywords: vbaof11.chm283000
f1_keywords:
- vbaof11.chm283000
ms.prod: office
api_name:
- Office.WorkflowTemplates
ms.assetid: 01df4716-4440-7761-8504-22f78e40f8e4
ms.date: 06/08/2017
---


# WorkflowTemplates Object (Office)

Represents a collection of  **WorkflowTemplate** objects.


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


## Properties



|**Name**|
|:-----|
|[Application](workflowtemplates-application-property-office.md)|
|[Count](workflowtemplates-count-property-office.md)|
|[Creator](workflowtemplates-creator-property-office.md)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
