---
title: Application.Options Property (Publisher)
keywords: vbapb10.chm131095
f1_keywords:
- vbapb10.chm131095
ms.prod: publisher
api_name:
- Publisher.Application.Options
ms.assetid: 999f208a-02e6-49fb-c9a0-42aa97c5e37e
ms.date: 06/08/2017
---


# Application.Options Property (Publisher)

Returns an  **[Options](options-object-publisher.md)** object that represents application settings you can set in Microsoft Publisher.


## Syntax

 _expression_. **Options**

 _expression_A variable that represents an  **Application** object.


### Return Value

Options


## Example

This example disables background saves and then saves the active publication.


```vb
Sub SetGlobalSaveOptions() 
 
 With Options 
 .AllowBackgroundSave = False 
 End With 
 
 ActiveDocument.Save 
 
End Sub
```


## See also


#### Concepts


 [Application Object](application-object-publisher.md)

