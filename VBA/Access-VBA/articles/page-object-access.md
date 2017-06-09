---
title: Page Object (Access)
keywords: vbaac10.chm10124
f1_keywords:
- vbaac10.chm10124
ms.prod: access
api_name:
- Access.Page
ms.assetid: 6351b0ea-bd07-5ee6-ea20-0d410e09d939
ms.date: 06/08/2017
---


# Page Object (Access)

A  **Page** object corresponds to an individual page on a tab control.


## Remarks

A  **Page** object is a member of a tab control's **[Pages](pages-object-access.md)** collection.

To return a reference to a particular  **Page** object in the **Pages** collection, use any of the following syntax forms.



|**Syntax**|**Description**|
|:-----|:-----|
|**Pages** ! _pagename_|The  _pagename_ argument is the name of the **Page** object.|
|**Pages** (" _pagename_")|The  _pagename_ argument is the name of the **Page** object.|
|**Pages** ( _index_)|The  _index_ argument is the numeric position of the object within the collection.|
You can create, move, or delete  **Page** objects and set their properties either in Visual Basic or in form Design view. To create a new **Page** object in Visual Basic, use the **Add** method of the **Pages** collection. To delete a **Page** object, use the **Remove** method of the **Pages** collection.

To create a new  **Page** object in form Design view, right-click the tab control and then click **Insert Page** on the shortcut menu. You can also copy an existing page and paste it. You can set the properties of the new **Page** object in form Design view by using the property sheet.

Each  **Page** object has a **PageIndex** property that indicates its position within the **Pages** collection. The **Value** property of the tab control is equal to the **PageIndex** property of the current page. You can use these properties to determine which page is currently selected after the user has switched from one page to another, or to change the order in which the pages appear in the control.

A  **Page** object is also a type of **Control** object. The **ControlType** property constant for a **Page** object is **acPage**. Although it is a control, a **Page** object belongs to a **Pages** collection, rather than a **Controls** collection. A tab control's **Pages** collection is a special type of **Controls** collection.

Each  **Page** object can also contain one or more controls. Controls on a **Page** object belong to that **Page** object's **Controls** collection. In order to work with a control on a **Page** object, you must refer to that control within the **Page** object's **Controls** collection.


## Events



|**Name**|
|:-----|
|[Click](page-click-event-access.md)|
|[DblClick](page-dblclick-event-access.md)|
|[MouseDown](page-mousedown-event-access.md)|
|[MouseMove](page-mousemove-event-access.md)|
|[MouseUp](page-mouseup-event-access.md)|

## Methods



|**Name**|
|:-----|
|[Move](page-move-method-access.md)|
|[Requery](page-requery-method-access.md)|
|[SetFocus](page-setfocus-method-access.md)|
|[SetTabOrder](page-settaborder-method-access.md)|
|[SizeToFit](page-sizetofit-method-access.md)|

## Properties



|**Name**|
|:-----|
|[Application](page-application-property-access.md)|
|[Caption](page-caption-property-access.md)|
|[Controls](page-controls-property-access.md)|
|[ControlTipText](page-controltiptext-property-access.md)|
|[ControlType](page-controltype-property-access.md)|
|[Enabled](page-enabled-property-access.md)|
|[EventProcPrefix](page-eventprocprefix-property-access.md)|
|[Height](page-height-property-access.md)|
|[HelpContextId](page-helpcontextid-property-access.md)|
|[InSelection](page-inselection-property-access.md)|
|[IsVisible](page-isvisible-property-access.md)|
|[Left](page-left-property-access.md)|
|[Name](page-name-property-access.md)|
|[OnClick](page-onclick-property-access.md)|
|[OnDblClick](page-ondblclick-property-access.md)|
|[OnMouseDown](page-onmousedown-property-access.md)|
|[OnMouseMove](page-onmousemove-property-access.md)|
|[OnMouseUp](page-onmouseup-property-access.md)|
|[PageIndex](page-pageindex-property-access.md)|
|[Parent](page-parent-property-access.md)|
|[Picture](page-picture-property-access.md)|
|[PictureData](page-picturedata-property-access.md)|
|[PictureType](page-picturetype-property-access.md)|
|[Properties](page-properties-property-access.md)|
|[Section](page-section-property-access.md)|
|[ShortcutMenuBar](page-shortcutmenubar-property-access.md)|
|[StatusBarText](page-statusbartext-property-access.md)|
|[Tag](page-tag-property-access.md)|
|[Top](page-top-property-access.md)|
|[Visible](page-visible-property-access.md)|
|[Width](page-width-property-access.md)|

## See also


#### Other resources


[Access Object Model Reference](http://msdn.microsoft.com/library/2de134a4-6c5c-d2a3-8377-f4dd973ba650%28Office.15%29.aspx)
