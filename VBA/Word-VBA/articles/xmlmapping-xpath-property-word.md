---
title: XMLMapping.XPath Property (Word)
keywords: vbawd10.chm199688198
f1_keywords:
- vbawd10.chm199688198
ms.prod: word
api_name:
- Word.XMLMapping.XPath
ms.assetid: 131234f2-ea3c-5b67-d10d-27c08aa94101
ms.date: 06/08/2017
---


# XMLMapping.XPath Property (Word)

Returns a  **String** that represents the XPath for the XML mapping, which evaluates to the currently mapped XML node. Read-only.


## Syntax

 _expression_ . **XPath**

 _expression_ An expression that returns an **[XMLMapping](xmlmapping-object-word.md)** object.


## Remarks

To set mapping for a content control, use the  **[SetMapping](xmlmapping-setmapping-method-word.md)** method or the **[SetMappingByNode](xmlmapping-setmappingbynode-method-word.md)** method. If the mapping is not active, using this property returns an error.


## Example

The following example checks whether the first content control in the active document is a date control and whether the XPath string is set to a specific built-in document property. It then sets the mapping to the control, if the XPath does not match and the control is a date control.


```vb
Dim objCC As ContentControl 
Dim objMap As XMLMapping 
Dim blnMap As Boolean 
 
Set objCC = ActiveDocument.ContentControls(1) 
Set objMap = objCC.XMLMapping 
 
If (objCC.Type = wdContentControlDate) And (objMap.XPath <> _ 
 "/ns1:coreProperties[1]/ns0:createdate[1]") Then 
 blnMap = objMap.SetMapping(XPath:="/ns1:coreProperties[1]/ns0:createdate[1]") 
 
 If blnMap = False Then 
 MsgBox "Unable to map the content control." 
 End If 
End If
```


## See also


#### Concepts


[XMLMapping Object](xmlmapping-object-word.md)

