---
title: Font.StylisticSets Property (Publisher)
keywords: vbapb10.chm5374016
f1_keywords:
- vbapb10.chm5374016
ms.prod: publisher
api_name:
- Publisher.Font.StylisticSets
ms.assetid: 0d25fbf3-8d68-c10f-0d1b-526314700329
ms.date: 06/08/2017
---


# Font.StylisticSets Property (Publisher)

Returns or sets a  **Variant** that represents the state of the **StylisticSets** property on the characters in a text range. Read/write.


## Syntax

 _expression_. **StylisticSets**

 _expression_A variable that represents a  **[Font](font-object-publisher.md)** object.


## Remarks

The  **StylisticSets** property applies from one to twenty increasingly complex sets of typography styles to the selected font.

Possible values for the  **StylisticSets** property and how they correspond to identifiers for stylistic sets in the user interface (UI) are shown in the following table. A value of zero (0) indicates that no stylistic set is applied.



|**StylisticSets property value**|**Stylistic set identifier in UI**|
|:-----|:-----|
|0|0|
|1|1|
|2|2|
|4|3|
|8|4|
The number of stylistic sets available varies, depending on the font.


 **Note**  The  **StylisticSets** property has an effect only for OpenType fonts that contain stylistic sets.


