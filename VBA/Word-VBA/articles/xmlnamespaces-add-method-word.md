---
title: XMLNamespaces.Add Method (Word)
keywords: vbawd10.chm248971365
f1_keywords:
- vbawd10.chm248971365
ms.prod: word
ms.assetid: 2b70fb44-adf0-31e9-0528-bda1189b85f5
ms.date: 06/08/2017
---


# XMLNamespaces.Add Method (Word)

 Returns an **XMLNamespace** object that represents a schema that is added to the Schema Library and made available to users in Microsoft Word.


## Syntax

 _expression_ . **Add**( **_Path_** , **_NamespaceURI_** , **_Alias_** , **_InstallForAllUsers_** )

 _expression_ Required. A variable that represents a **** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Path_|Required| **String**|The path and file name of the schema. This may be a local file path, a network path, or an Internet address.|
| _NamespaceURI_|Optional| **String**|The namespace Uniform Resource Indicator as specified in the schema. The NamespaceURI parameter is case-sensitive and must be spelled exactly as specified in schema.|
| _Alias_|Optional| **String**|The name of the schema as it appears on the  **Schemas** tab in the **Templates and Add-ins** dialog box.|
| _InstallForAllUsers_|Optional| **Boolean**| **True** if all users that log on to a computer can access and use the new schema. The default is **False** .|

### Return Value

XMLNamespace


## Example

The following example adds the specified schema to the Schema Library and then attaches it to the active document. This example assumes that you have a schema named simplesample.xsd at the specified path.


```vb
Sub AddSchema() 
 Dim objSchema As XMLNamespace 
 
 Set objSchema = Application.XMLNamespaces _ 
 .Add ("c:\schemas\simplesample.xsd") 
 
 objSchema.AttachToDocument ActiveDocument 
End Sub
```


## See also


#### Concepts




