---
title: OLEFormat.DoVerb Method (Publisher)
keywords: vbapb10.chm4456455
f1_keywords:
- vbapb10.chm4456455
ms.prod: publisher
api_name:
- Publisher.OLEFormat.DoVerb
ms.assetid: c4bca1f2-a3dd-0c49-1268-40e68e1fcef0
ms.date: 06/08/2017
---


# OLEFormat.DoVerb Method (Publisher)

Requests that an OLE object perform one of its verbs.


## Syntax

 _expression_. **DoVerb**( **_iVerb_**)

 _expression_A variable that represents an  **OLEFormat** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|iVerb|Required| **Long**|The verb to perform. |

## Remarks

Use the  **[ObjectVerbs](oleformat-objectverbs-property-publisher.md)** property to determine the available verbs for an OLE object.


## Example

This example performs the first verb for the third shape on the first page of the active publication if the shape is a linked or embedded OLE object.


```vb
With ActiveDocument.Pages(1).Shapes(3) 
 If .Type = pbEmbeddedOLEObject Or _ 
 .Type = pbLinkedOLEObject Then 
 .OLEFormat.DoVerb (1) 
 End If 
End With
```

This example performs the verb "Open" for the third shape on the first page of the active publication if the shape is an OLE object that supports the verb "Open."




```vb
Dim strVerb As String 
Dim intVerb As Integer 
 
With ActiveDocument.Pages(1).Shapes(3) 
 
 ' Verify that the shape is an OLE object. 
 If .Type = pbEmbeddedOLEObject Or _ 
 .Type = pbLinkedOLEObject Then 
 
 ' Loop through the ObjectVerbs collection 
 ' until the "Open" verb is found. 
 For Each strVerb In .OLEFormat.ObjectVerbs 
 intVerb = intVerb + 1 
 If strVerb = "Open" Then 
 
 ' Perform the "Open" verb. 
 .OLEFormat.DoVerb iVerb:=intVerb 
 Exit For 
 End If 
 Next strVerb 
 End If 
End With 

```


