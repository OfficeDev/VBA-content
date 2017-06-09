---
title: Hyperlink.CreateNewDocument Method (Word)
keywords: vbawd10.chm161284202
f1_keywords:
- vbawd10.chm161284202
ms.prod: word
api_name:
- Word.Hyperlink.CreateNewDocument
ms.assetid: e3077a0d-6a83-e36d-7199-8ec6aca8dfa7
ms.date: 06/08/2017
---


# Hyperlink.CreateNewDocument Method (Word)

Creates a new document linked to the specified hyperlink.


## Syntax

 _expression_ . **CreateNewDocument**( **_FileName_** , **_EditNow_** , **_Overwrite_** )

 _expression_ Required. A variable that represents a **[Hyperlink](hyperlink-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FileName_|Required| **String**|The file name of the specified document.|
| _EditNow_|Required| **Boolean**| **True** to have the specified document open immediately in its associated editing environment. The default value is **True** .|
| _Overwrite_|Required| **Boolean**| **True** to overwrite any existing file of the same name in the same folder. **False** if any existing file of the same name is preserved and the FileName argument specifies a new file name. The default value is **False** .|

## Example

This example creates a new document based on the new hyperlink in the first document and then loads the new document into Microsoft Word for editing. The document is called ?Overview.doc,? and it overwrites any file of the same name in the  `\\Server1\Annual` folder.


```vb
With Documents(1) 
 Set objHyper = _ 
 .Hyperlinks.Add(Anchor:=Selection.Range, _ 
 Address:="\\Server1\Annual\Overview.doc") 
 objHyper.CreateNewDocument _ 
 FileName:="\\Server1\Annual\Overview.doc", _ 
 EditNow:=True, Overwrite:=True 
End With
```


## See also


#### Concepts


[Hyperlink Object](hyperlink-object-word.md)

