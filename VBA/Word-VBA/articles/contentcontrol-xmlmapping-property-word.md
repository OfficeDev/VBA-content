---
title: ContentControl.XMLMapping Property (Word)
keywords: vbawd10.chm266534916
f1_keywords:
- vbawd10.chm266534916
ms.prod: word
api_name:
- Word.ContentControl.XMLMapping
ms.assetid: 3730e4b2-b69c-3428-6968-4a48a3dc0b93
ms.date: 06/08/2017
---


# ContentControl.XMLMapping Property (Word)

Returns an  **[XMLMapping](xmlmapping-object-word.md)** object that represents the mapping of a content control to XML data in the data store of a document. Read-only.


## Syntax

 _expression_ . **XMLMapping**

 _expression_ An expression that returns a **ContentControl** object.


## Example

The following example sets the built-in Author document property and adds a new content control to the active document, and then sets the mapping for the control to the value of the built-in document property.


```vb
Dim objCC As ContentControl 
Dim objMap As XMLMapping 
Dim blnMap As Boolean 
 
ActiveDocument.BuiltInDocumentProperties("Author").Value = "David Jaffe" 
 
Set objCC = ActiveDocument.ContentControls.Add _ 
 (wdContentControlText, ActiveDocument.Paragraphs(1).Range) 
 
Set objMap = objCC.XMLMapping 
blnMap = objMap.SetMapping(XPath:="/ns1:coreProperties[1]/ns0:creator[1]") 
 
If blnMap = False Then 
 MsgBox "Unable to map the content control." 
End If
```


## See also


#### Concepts


[ContentControl Object](contentcontrol-object-word.md)

