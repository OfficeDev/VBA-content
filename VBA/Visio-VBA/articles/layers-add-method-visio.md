---
title: Layers.Add Method (Visio)
keywords: vis_sdr.chm11916675
f1_keywords:
- vis_sdr.chm11916675
ms.prod: visio
api_name:
- Visio.Layers.Add
ms.assetid: e46bc30f-ad35-ddeb-86d3-14ef535451cf
ms.date: 06/08/2017
---


# Layers.Add Method (Visio)

Adds a new  **Layer** object to a **Layers** collection.


## Syntax

 _expression_ . **Add**( **_LayerName_** )

 _expression_ A variable that represents a **Layers** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _LayerName_|Required| **String**|The name of the new layer.|

### Return Value

Layer


## Example

The following macro shows how to add a  **Layer** object to the **Layers** collection.


```vb
Public Sub AddLayer_Example() 
 
 Dim vsoDocument As Visio.Document 
 Dim vsoPages As Visio.Pages 
 Dim vsoPage As Visio.Page 
 Dim vsoLayers As Visio.Layers 
 Dim vsoLayer As Visio.Layer 
 
 'Add a document based on the Basic Diagram template. 
 Set vsoDocument = Documents.Add("Basic Diagram.vst") 
 
 'Get the Pages collection and add a page to the collection. 
 Set vsoPages = vsoDocument.Pages 
 Set vsoPage = vsoPages.Add 
 
 'Get the Layers collection and add a layer named "MyLayer" 
 'to the collection. 
 Set vsoLayers = vsoPage.Layers 
 Set vsoLayer = vsoLayers.Add("MyLayer") 
 
End Sub
```


