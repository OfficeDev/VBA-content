---
title: OLEFormat Object (Word)
keywords: vbawd10.chm2355
f1_keywords:
- vbawd10.chm2355
ms.prod: word
api_name:
- Word.OLEFormat
ms.assetid: d4c7aa65-5d3a-0b79-914b-6f908b506f63
ms.date: 06/08/2017
---


# OLEFormat Object (Word)

Represents the OLE characteristics (other than linking) for an OLE object, ActiveX control, or field.


## Remarks

Use the  **OLEFormat** property for a shape, inline shape, or field to return the **OLEFormat** object. The following example displays the class type for the first shape on the active document.


```
MsgBox ActiveDocument.Shapes(1).OLEFormat.ClassType
```

Not all types of shapes, inline shapes, and fields have OLE capabilities. Use the  **Type** property for the **Shape** and **InlineShape** objects to determine what category the specified shape or inline shape falls into. The **Type** property for a **Field** object returns the type of field.

You can use the  **Activate**, **Edit**, **Open**, and **DoVerb** methods to automate an OLE object.

Use the  **Object** property to return an object that represents an ActiveX control or OLE object. With this object, you can use the properties and methods of the container application or the ActiveX control.


## Methods



|**Name**|
|:-----|
|[Activate](oleformat-activate-method-word.md)|
|[ActivateAs](oleformat-activateas-method-word.md)|
|[ConvertTo](oleformat-convertto-method-word.md)|
|[DoVerb](oleformat-doverb-method-word.md)|
|[Edit](oleformat-edit-method-word.md)|
|[Open](oleformat-open-method-word.md)|

## Properties



|**Name**|
|:-----|
|[Application](oleformat-application-property-word.md)|
|[ClassType](oleformat-classtype-property-word.md)|
|[Creator](oleformat-creator-property-word.md)|
|[DisplayAsIcon](oleformat-displayasicon-property-word.md)|
|[IconIndex](oleformat-iconindex-property-word.md)|
|[IconLabel](oleformat-iconlabel-property-word.md)|
|[IconName](oleformat-iconname-property-word.md)|
|[IconPath](oleformat-iconpath-property-word.md)|
|[Label](oleformat-label-property-word.md)|
|[Object](oleformat-object-property-word.md)|
|[Parent](oleformat-parent-property-word.md)|
|[PreserveFormattingOnUpdate](oleformat-preserveformattingonupdate-property-word.md)|
|[ProgID](oleformat-progid-property-word.md)|

## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)
