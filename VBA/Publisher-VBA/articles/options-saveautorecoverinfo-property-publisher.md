---
title: Options.SaveAutoRecoverInfo Property (Publisher)
keywords: vbapb10.chm1048599
f1_keywords:
- vbapb10.chm1048599
ms.prod: publisher
api_name:
- Publisher.Options.SaveAutoRecoverInfo
ms.assetid: 1cbb7960-8995-37f4-5989-01b97152269f
ms.date: 06/08/2017
---


# Options.SaveAutoRecoverInfo Property (Publisher)

 **True** if Microsoft Publisher automatically saves publications for recovery if the application is unexpectedly shut down. Read/write **Boolean**.


## Syntax

 _expression_. **SaveAutoRecoverInfo**

 _expression_A variable that represents a  **Options** object.


### Return Value

Boolean


## Remarks

Use the  **[SaveAutoRecoverInfoInterval](options-saveautorecoverinfointerval-property-publisher.md)** property to specify how often auto recovery saves occur.


## Example

This example enables the global auto recovery option and sets the save interval to every five minutes.


```vb
Sub SetAutoRecoverInfo() 
 With Options 
 .SaveAutoRecoverInfo = True 
 .SaveAutoRecoverInfoInterval = 5 
 End With 
End Sub
```


