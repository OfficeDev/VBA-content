---
title: CardView Object (Outlook)
keywords: vbaol11.chm3207
f1_keywords:
- vbaol11.chm3207
ms.prod: outlook
api_name:
- Outlook.CardView
ms.assetid: cdac229b-f2b6-9ecb-e1a7-b53509426570
ms.date: 06/08/2017
---


# CardView Object (Outlook)

Represents a view that displays Outlook items as a series of index cards.


## Remarks

The  **CardView** object, derived from the **[View](view-object-outlook.md)** object, allows you to create customizable views that allow you to display Outlook items as index cards.

Outlook provides several built-in  **CardView** objects, and you can also create custom **CardView** objects. Use the **[Add](views-add-method-outlook.md)** method of the **[Views](views-object-outlook.md)** collection to add a new **CardView** to a **[Folder](folder-object-outlook.md)** object. Use the **[Standard](cardview-standard-property-outlook.md)** property to determine if an existing **CardView** object is built-in or custom.

You can configure how Outlook items appear within the  **CardView** object. Use the **[MultiLineFieldHeight](cardview-multilinefieldheight-property-outlook.md)** property to specify the number of lines used to display multi-line text in each card, the **[HeadingsFont](cardview-headingsfont-property-outlook.md)** property to specify the font used to display heading text on each card, and the **[BodyFont](cardview-bodyfont-property-outlook.md)** property to specify the font used to display body text on each card. Use the **[AllowInCellEditing](cardview-allowincellediting-property-outlook.md)** property to allow editing of Outlook item property values in the view, and the **[ShowEmptyFields](cardview-showemptyfields-property-outlook.md)** property to display empty Outlook item properties in the view. Use the **[Filter](cardview-filter-property-outlook.md)** property to determine which Outlook items to display in the view, the **[ViewFields](cardview-viewfields-property-outlook.md)** collection to specify the Outlook item properties to display in each card, and the **[SortFields](cardview-sortfields-property-outlook.md)** collection to specify the Outlook item properties by which Outlook items are sorted in the view.

The definition for each  **CardView** object is stored in Extensible Markup Language (XML) format. Use the **[XML](cardview-xml-property-outlook.md)** property to work with the XML definition for the **CardView** object.

Use the  **[Apply](cardview-apply-method-outlook.md)** method to apply any changes made to the **CardView** object to the current view. Use the **[Save](cardview-save-method-outlook.md)** method to persist any changes made to the **CardView** object. Use the **[LockUserChanges](cardview-lockuserchanges-property-outlook.md)** property to allow or prevent changes to the user interface for the view.

You can change built-in  **CardView** objects, but you cannot delete them. Use the **[Delete](cardview-delete-method-outlook.md)** method to delete a custom **CardView** object. Use the **[Reset](cardview-reset-method-outlook.md)** method to reset the properties of a built-in **CardView** object to their default values.


## Methods



|**Name**|
|:-----|
|[Apply](cardview-apply-method-outlook.md)|
|[Copy](cardview-copy-method-outlook.md)|
|[Delete](cardview-delete-method-outlook.md)|
|[GoToDate](cardview-gotodate-method-outlook.md)|
|[Reset](cardview-reset-method-outlook.md)|
|[Save](cardview-save-method-outlook.md)|

## Properties



|**Name**|
|:-----|
|[AllowInCellEditing](cardview-allowincellediting-property-outlook.md)|
|[Application](cardview-application-property-outlook.md)|
|[AutoFormatRules](cardview-autoformatrules-property-outlook.md)|
|[BodyFont](cardview-bodyfont-property-outlook.md)|
|[Class](cardview-class-property-outlook.md)|
|[Filter](cardview-filter-property-outlook.md)|
|[HeadingsFont](cardview-headingsfont-property-outlook.md)|
|[Language](cardview-language-property-outlook.md)|
|[LockUserChanges](cardview-lockuserchanges-property-outlook.md)|
|[MultiLineFieldHeight](cardview-multilinefieldheight-property-outlook.md)|
|[Name](cardview-name-property-outlook.md)|
|[Parent](cardview-parent-property-outlook.md)|
|[SaveOption](cardview-saveoption-property-outlook.md)|
|[Session](cardview-session-property-outlook.md)|
|[ShowEmptyFields](cardview-showemptyfields-property-outlook.md)|
|[SortFields](cardview-sortfields-property-outlook.md)|
|[Standard](cardview-standard-property-outlook.md)|
|[ViewFields](cardview-viewfields-property-outlook.md)|
|[ViewType](cardview-viewtype-property-outlook.md)|
|[Width](cardview-width-property-outlook.md)|
|[XML](cardview-xml-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
