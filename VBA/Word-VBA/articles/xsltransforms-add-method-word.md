---
title: XSLTransforms.Add Method (Word)
keywords: vbawd10.chm99221605
f1_keywords:
- vbawd10.chm99221605
ms.prod: word
ms.assetid: 017e0389-c414-3c73-4b9f-a130982339d2
ms.date: 06/08/2017
---


# XSLTransforms.Add Method (Word)

Returns an  **XSLTransform** object that represents an Extensible Stylesheet Language Transformation (XSLT) added to the collection of XSLTs for a specified schema.


## Syntax

 _expression_ . **Add**( **_Location_** , **_Alias_** , **_InstallForAllUsers_** )

 _expression_ Required. A variable that represents a **** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Location_|Required| **String**|The path and file name of the XSLT. This may be a local file path, a network path, or an Internet address.|
| _Alias_|Optional| **String**|The name of the XSLT as it appears in the Schema Library.|
| _InstallForAllUsers_|Optional| **Boolean**| **True** if all users that log on to a computer can access and use the new schema. The default is **False** .|

### Return Value

XSLTransform


## Example

The following code adds a schema to the Schema Library and then adds an XSLT to the newly added schema.


```vb
Sub AddXSLT() 
 Dim objSchema As XMLNamespace 
 Dim objTransform As XSLTransform 
 
 Set objSchema = Application.XMLNamespaces("SimpleSample") 
 Set objTransform = objSchema.XSLTransforms _ 
 .Add("c:\schemas\simplesample.xsl") 
 
End Sub
```


## See also


#### Concepts




