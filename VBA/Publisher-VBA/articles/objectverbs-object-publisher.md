---
title: ObjectVerbs Object (Publisher)
keywords: vbapb10.chm4587519
f1_keywords:
- vbapb10.chm4587519
ms.prod: publisher
api_name:
- Publisher.ObjectVerbs
ms.assetid: e04cf7db-ee56-7d95-9f5c-7ecee1844866
ms.date: 06/08/2017
---


# ObjectVerbs Object (Publisher)

Represents the collection of OLE verbs for the specified OLE object. OLE verbs are the operations supported by an OLE object. Commonly used OLE verbs are play and edit.
 


## Example

Use the  **[ObjectVerbs](oleformat-objectverbs-property-publisher.md)** property to return an **ObjectVerbs** object. The following example displays all the available verbs for the OLE object contained in the first shape on first page in the active publication. For this example to work, the specified shape must contain an OLE object.
 

 

```
Sub GetVerbs() 
 Dim intCount As Integer 
 
 With ActiveDocument.Pages(1).Shapes(1).OLEFormat 
 For intCount = 1 To .ObjectVerbs.Count 
 MsgBox .ObjectVerbs(intCount) 
 Next 
 End With 
End Sub
```


## Methods



|**Name**|
|:-----|
|[Item](objectverbs-item-method-publisher.md)|

## Properties



|**Name**|
|:-----|
|[Application](objectverbs-application-property-publisher.md)|
|[Count](objectverbs-count-property-publisher.md)|
|[Parent](objectverbs-parent-property-publisher.md)|

