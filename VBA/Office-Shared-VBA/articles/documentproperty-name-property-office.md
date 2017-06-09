---
title: DocumentProperty.Name Property (Office)
keywords: vbaof11.chm250005
f1_keywords:
- vbaof11.chm250005
ms.prod: office
api_name:
- Office.DocumentProperty.Name
ms.assetid: b609c38e-71ca-e019-9852-fc7811dc798f
ms.date: 06/08/2017
---


# DocumentProperty.Name Property (Office)

Gets or sets the name of a document property. Read/write.


## Syntax

 _expression_. **Name**( **_lcid_**, **_pbstrRetVal_** )

 _expression_ A variable that represents a **DocumentProperty** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _lcid_|Required|**Long**|Represents the language identifier.|
| _pbstrRetVal_|Required|**String**|Represents the return value for the property.|

### Return Value

String


## Remarks

A  **DocumentProperty** object represents a custom or built-in document property of a container document.


## Example

This example displays the name, type, and value of a document property. You must pass a valid  **DocumentProperty** object to the procedure.


```
Sub DisplayPropertyInfo(dp As DocumentProperty) 
 MsgBox "value = " &amp; dp.Value &amp; Chr(13) &amp; _ 
 "type = " &amp; dp.Type &amp; Chr(13) &amp; _ 
 "name = " &amp; dp.Name 
End Sub
```


## See also


#### Concepts


[DocumentProperty Object](documentproperty-object-office.md)
#### Other resources


[DocumentProperty Object Members](documentproperty-members-office.md)

