---
title: ListFormat Object (Word)
keywords: vbawd10.chm2496
f1_keywords:
- vbawd10.chm2496
ms.prod: word
api_name:
- Word.ListFormat
ms.assetid: 74773fd6-b713-34d4-b7be-f543c983008d
ms.date: 06/08/2017
---


# ListFormat Object (Word)

Represents the list formatting attributes that can be applied to the paragraphs in a range.


## Remarks

Use the  **ListFormat** property to return the **ListFormat** object for a range. The following example applies the default bulleted list format to the selection.


```
Selection.Range.ListFormat.ApplyBulletDefault
```

An easy way to apply list formatting is to use the  **ApplyBulletDefault**, **ApplyNumberDefault**, and **ApplyOutlineNumberDefault** methods, which correspond, respectively, to the first list format (excluding **None**) on each tab in the  **Bullets and Numbering** dialog box.

To apply a format other than the default format, use the  **ApplyListTemplate** method, which allows you to specify the list format (list template) you want to apply.

Use the  **List** or **ListTemplate** property to return the list or list template from the first paragraph in the specified range.

Use the  **ListFormat** property with a **Range** object to access the list formatting properties and methods available for the specified range. The following example applies the default bulleted list format to the second paragraph in the active document.




```
ActiveDocument.Paragraphs(2).Range.ListFormat.ApplyBulletDefault
```

However, if there is already a list defined in your document, you can access a  **List** object by using the **Lists** property. The following example changes the format of the list created in the preceding example to the first number format on the **Numbered** tab in the **Bullets and Numbering** dialog box.




```
ActiveDocument.Lists(1).ApplyListTemplate _ 
 ListTemplate:=ListGalleries(2).ListTemplates(1)
```


## Methods



|**Name**|
|:-----|
|[ApplyBulletDefault](listformat-applybulletdefault-method-word.md)|
|[ApplyListTemplate](listformat-applylisttemplate-method-word.md)|
|[ApplyListTemplateWithLevel](listformat-applylisttemplatewithlevel-method-word.md)|
|[ApplyNumberDefault](listformat-applynumberdefault-method-word.md)|
|[ApplyOutlineNumberDefault](listformat-applyoutlinenumberdefault-method-word.md)|
|[CanContinuePreviousList](listformat-cancontinuepreviouslist-method-word.md)|
|[ConvertNumbersToText](listformat-convertnumberstotext-method-word.md)|
|[CountNumberedItems](listformat-countnumbereditems-method-word.md)|
|[ListIndent](listformat-listindent-method-word.md)|
|[ListOutdent](listformat-listoutdent-method-word.md)|
|[RemoveNumbers](listformat-removenumbers-method-word.md)|

## Properties



|**Name**|
|:-----|
|[Application](listformat-application-property-word.md)|
|[Creator](listformat-creator-property-word.md)|
|[List](listformat-list-property-word.md)|
|[ListLevelNumber](listformat-listlevelnumber-property-word.md)|
|[ListPictureBullet](listformat-listpicturebullet-property-word.md)|
|[ListString](listformat-liststring-property-word.md)|
|[ListTemplate](listformat-listtemplate-property-word.md)|
|[ListType](listformat-listtype-property-word.md)|
|[ListValue](listformat-listvalue-property-word.md)|
|[Parent](listformat-parent-property-word.md)|
|[SingleList](listformat-singlelist-property-word.md)|
|[SingleListTemplate](listformat-singlelisttemplate-property-word.md)|

## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)
