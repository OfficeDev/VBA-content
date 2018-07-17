---
title: TableView Object (Outlook)
keywords: vbaol11.chm3204
f1_keywords:
- vbaol11.chm3204
ms.prod: outlook
api_name:
- Outlook.TableView
ms.assetid: 026e27f8-1655-060d-e8cc-87eaaf4f1510
ms.date: 06/08/2017
---


# TableView Object (Outlook)

Represents a view that displays Outlook items in a table, with each item in a row and the details of the item in the columns.


## Remarks

The  **TableView** object, derived from the **[View](view-object-outlook.md)** object, allows you to create customizable views that allow you to display Outlook items in a table.

Outlook provides several built-in  **TableView** objects, and you can also create custom **TableView** objects. Use the **[Add](views-add-method-outlook.md)** method of the **[Views](views-object-outlook.md)** collection to add a new **TableView** to a **[Folder](folder-object-outlook.md)** object. Use the **Standard** property to determine if an existing **TableView** object is built-in or custom.

You can configure the appearance and functionality of the  **TableView** object. Use the **[AutomaticColumnSizing](tableview-automaticcolumnsizing-property-outlook.md)** property to determine whether the view automatically resizes columns and the **[AutomaticGrouping](tableview-automaticgrouping-property-outlook.md)** property to determine if the view automatically groups Outlook items. Use the **[AutoPreview](tableview-autopreview-property-outlook.md)** property to determine whether preview information is displayed within the row for an Outlook item in the view, and the **[AutoPreviewFont](tableview-autopreviewfont-property-outlook.md)** property to specify the font used to display preview information. Use the **[Multiline](tableview-multiline-property-outlook.md)** property to determine whether to show Outlook items in multiline mode.

You can also configure how Outlook items appear within the  **TableView** object. Use the **[ColumnFont](tableview-columnfont-property-outlook.md)** property to specify the font used for column headers and the **[RowFont](tableview-rowfont-property-outlook.md)** property to specify the font used for Outlook items in the view. Use the **[AllowInCellEditing](tableview-allowincellediting-property-outlook.md)** property to allow editing of Outlook item property values in the view. Use the **[Filter](tableview-filter-property-outlook.md)** property to determine which Outlook items to display in the view and the **[ViewFields](tableview-viewfields-property-outlook.md)** collection to specify the Outlook item properties to display for each Outlook item. Use the **[GroupByFields](tableview-groupbyfields-property-outlook.md)** to specify the Outlook item properties by which Outlook items are grouped, and the **[SortFields](tableview-sortfields-property-outlook.md)** collection to specify the Outlook item properties by which Outlook items are sorted in the view.

The definition for each  **TableView** object is stored in Extensible Markup Language (XML) format. Use the **[XML](tableview-xml-property-outlook.md)** property to work with the XML definition for the **TableView** object.

Use the  **[Apply](tableview-apply-method-outlook.md)** method to apply any changes made to the **TableView** object to the current view. Use the **[Save](tableview-save-method-outlook.md)** method to persist any changes made to the **TableView** object. Use the **[LockUserChanges](tableview-lockuserchanges-property-outlook.md)** property to allow or prevent changes to the user interface for the view.

You can change built-in  **TableView** objects, but you cannot delete them. Use the **[Delete](tableview-delete-method-outlook.md)** method to delete a custom **TableView** object. Use the **[Reset](tableview-reset-method-outlook.md)** method to reset the properties of a built-in **TableView** object to their default values.


## Methods



|**Name**|
|:-----|
|[Apply](tableview-apply-method-outlook.md)|
|[Copy](tableview-copy-method-outlook.md)|
|[Delete](tableview-delete-method-outlook.md)|
|[GetTable](tableview-gettable-method-outlook.md)|
|[GoToDate](tableview-gotodate-method-outlook.md)|
|[Reset](tableview-reset-method-outlook.md)|
|[Save](tableview-save-method-outlook.md)|

## Properties



|**Name**|
|:-----|
|[AllowInCellEditing](tableview-allowincellediting-property-outlook.md)|
|[AlwaysExpandConversation](tableview-alwaysexpandconversation-property-outlook.md)|
|[Application](tableview-application-property-outlook.md)|
|[AutoFormatRules](tableview-autoformatrules-property-outlook.md)|
|[AutomaticColumnSizing](tableview-automaticcolumnsizing-property-outlook.md)|
|[AutomaticGrouping](tableview-automaticgrouping-property-outlook.md)|
|[AutoPreview](tableview-autopreview-property-outlook.md)|
|[AutoPreviewFont](tableview-autopreviewfont-property-outlook.md)|
|[Class](tableview-class-property-outlook.md)|
|[ColumnFont](tableview-columnfont-property-outlook.md)|
|[DefaultExpandCollapseSetting](tableview-defaultexpandcollapsesetting-property-outlook.md)|
|[Filter](tableview-filter-property-outlook.md)|
|[GridLineStyle](tableview-gridlinestyle-property-outlook.md)|
|[GroupByFields](tableview-groupbyfields-property-outlook.md)|
|[HideReadingPaneHeaderInfo](tableview-hidereadingpaneheaderinfo-property-outlook.md)|
|[Language](tableview-language-property-outlook.md)|
|[LockUserChanges](tableview-lockuserchanges-property-outlook.md)|
|[MaxLinesInMultiLineView](tableview-maxlinesinmultilineview-property-outlook.md)|
|[MultiLine](tableview-multiline-property-outlook.md)|
|[MultiLineWidth](tableview-multilinewidth-property-outlook.md)|
|[Name](tableview-name-property-outlook.md)|
|[Parent](tableview-parent-property-outlook.md)|
|[RowFont](tableview-rowfont-property-outlook.md)|
|[SaveOption](tableview-saveoption-property-outlook.md)|
|[Session](tableview-session-property-outlook.md)|
|[ShowConversationByDate](tableview-showconversationbydate-property-outlook.md)|
|[ShowConversationSendersAboveSubject](tableview-showconversationsendersabovesubject-property-outlook.md)|
|[ShowFullConversations](tableview-showfullconversations-property-outlook.md)|
|[ShowItemsInGroups](tableview-showitemsingroups-property-outlook.md)|
|[ShowNewItemRow](tableview-shownewitemrow-property-outlook.md)|
|[ShowReadingPane](tableview-showreadingpane-property-outlook.md)|
|[SortFields](tableview-sortfields-property-outlook.md)|
|[Standard](tableview-standard-property-outlook.md)|
|[ViewFields](tableview-viewfields-property-outlook.md)|
|[ViewType](tableview-viewtype-property-outlook.md)|
|[XML](tableview-xml-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
