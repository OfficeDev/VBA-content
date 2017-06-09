---
title: Research.FavoriteService Property (Word)
keywords: vbawd10.chm201655275
f1_keywords:
- vbawd10.chm201655275
ms.prod: word
api_name:
- Word.Research.FavoriteService
ms.assetid: ed8654bb-6f70-fe66-70cf-5736163028d4
ms.date: 06/08/2017
---


# Research.FavoriteService Property (Word)

Returns or sets a  **String** that specifies the favorite research service.


## Syntax

 _expression_ . **FavoriteService**

 _expression_ An expression that returns a **[Research](research-object-word.md)** object.


## Remarks

The  **String** that is set or returned for this property specifies the GUID of the favorite research service.

Setting this property has the same effect as choosing a favorite research service through the Research Options dialog in Word. 


 **Note**  The GUIDs for all installed research services can be located in the `HKCU\Software\Microsoft\Office\14.0\Common\Research\Sources` registry key.


## Example

The following code example changes the favorite research service to "Encarta Dictionary: English (North America)".


```vb
Dim objResearch As Research 
 
Sub MyFunction() 
 
Set objResearch = Research 
 
'Set the favorite service 
objResearch.FavoriteService = "FEF89077-4F4D-4803-A8BF-228083F70EAA" 
 
End Sub
```


## See also


#### Concepts


[Research Object](research-object-word.md)

