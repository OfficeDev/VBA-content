---
title: PhoneticGuide Object (Publisher)
keywords: vbapb10.chm6225919
f1_keywords:
- vbapb10.chm6225919
ms.prod: publisher
api_name:
- Publisher.PhoneticGuide
ms.assetid: 164e8b54-4bad-4de9-bf6e-52c5687dfbc6
ms.date: 06/08/2017
---


# PhoneticGuide Object (Publisher)

Represents base text with supplementary text appearing above it as a guide to pronunciation.
 


## Example

Use the  **PhoneticGuide** property of a **Field** object to return an existing **PhoneticGuide** object. Use the **AddPhoneticGuide** method of a **Fields** collection to create a new **PhoneticGuide** object.
 

 

 

 
The following example adds a new  **PhoneticGuide** object to the active publication.
 

 



```
Selection.TextRange.Fields.AddPhoneticGuide _ 
 Range:=Selection.TextRange, Text:="ver-E nIs", _ 
 Alignment:=pbPhoneticGuideAlignmentCenter, _ 
 Raise:=11, FontSize:=7
```


## Methods



|**Name**|
|:-----|
|[Clear](phoneticguide-clear-method-publisher.md)|

## Properties



|**Name**|
|:-----|
|[Alignment](phoneticguide-alignment-property-publisher.md)|
|[Application](phoneticguide-application-property-publisher.md)|
|[BaseText](phoneticguide-basetext-property-publisher.md)|
|[FontName](phoneticguide-fontname-property-publisher.md)|
|[FontSize](phoneticguide-fontsize-property-publisher.md)|
|[Parent](phoneticguide-parent-property-publisher.md)|
|[Raise](phoneticguide-raise-property-publisher.md)|
|[Text](phoneticguide-text-property-publisher.md)|

