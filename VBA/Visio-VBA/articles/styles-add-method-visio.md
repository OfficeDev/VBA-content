---
title: Styles.Add Method (Visio)
keywords: vis_sdr.chm11516680
f1_keywords:
- vis_sdr.chm11516680
ms.prod: visio
api_name:
- Visio.Styles.Add
ms.assetid: def0d922-048a-eab6-51cd-6052ba96fea8
ms.date: 06/08/2017
---


# Styles.Add Method (Visio)

Adds a new  **Style** object to a **Styles** collection.


## Syntax

 _expression_ . **Add**( **_StyleName_** , **_BasedOn_** , **_fIncludesText_** , **_fIncludesLine_** , **_fIncludesFill_** )

 _expression_ A variable that represents a **Styles** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _StyleName_|Required| **String**|The new style name.|
| _BasedOn_|Required| **String**|The name of the style on which to base the new style.|
| _fIncludesText_|Required| **Integer**|Zero to disable text attributes, or non-zero to enable them.|
| _fIncludesLine_|Required| **Integer**|Zero to disable line attributes, or non-zero to enable them.|
| _fIncludesFill_|Required| **Integer**|Zero to disable fill attributes, or non-zero to enable them.|

### Return Value

Style


## Remarks

To base the new style on no style, pass a zero-length string ("") for the  _BasedOn_ argument.


## Example

The following macro shows how to add  **Style** objects to the **Styles** collection. It shows how to add a new style based on an existing style, as well as how to add a new style created from scratch.


```vb
Public Sub AddStyle_Example() 
 
 Dim vsoDocument As Visio.Document 
 Dim vsoStyles As Visio.Styles 
 Dim vsoStyle As Visio.Style 
 
 'Add a document based on the Basic Diagram template. 
 Set vsoDocument = Documents.Add("Basic Diagram.vst") 
 
 'Add a style named "My FillStyle" to the Styles collection. 
 'This style is based on the Basic style and includes 
 'only a Fill style. 
 Set vsoStyles = vsoDocument.Styles 
 Set vsoStyle = vsoStyles.Add("My FillStyle", "Basic", 0, 0, 1) 
 
 'Add a style named "My NoStyle" to the Styles collection. 
 'This style is not based on an existing style and includes 
 'Text, Line, and Fill styles. 
 Set vsoStyle = vsoStyles.Add("My NoStyle", "", 1, 1, 1) 
 
End Sub
```


