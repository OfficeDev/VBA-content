---
title: NoteItem Object (Outlook)
keywords: vbaol11.chm3001
f1_keywords:
- vbaol11.chm3001
ms.prod: outlook
api_name:
- Outlook.NoteItem
ms.assetid: ddf5baaa-6e13-a6fb-96e8-311e7761fa98
ms.date: 06/08/2017
---


# NoteItem Object (Outlook)

Represents a note in a Notes folder.


## Remarks

A  **NoteItem** is not customizable. If you open a new note, you will notice that it is not possible to place it in design time.

The  **[Subject](noteitem-subject-property-outlook.md)** property of a **NoteItem** object is read-only because it is calculated from the body text of the note. Also, the **NoteItem** **[Body](noteitem-body-property-outlook.md)** can only be rich text, so the properties that correspond to HTML and Microsoft Word content do not apply. Although the **[GetInspector](noteitem-getinspector-property-outlook.md)** property will work on notes, because notes can't be customized, some of the **[Inspector](inspector-object-outlook.md)** properties, methods, and events will not apply to **NoteItem** objects.

Use the  **[CreateItem](application-createitem-method-outlook.md)** method to create a **NoteItem** object that represents a new note.

Use  **[Items](items-item-method-outlook.md)** ( _index_ ), where _index_ is the index number of a note or a value used to match the default property of a note, to return a single **NoteItem** object from a Notes folder.


## Example

 The following Microsoft Visual Basic example returns a new note.


```
Set myItem = Application.CreateItem(olNoteItem)
```


## Methods



|**Name**|
|:-----|
|[Close](noteitem-close-method-outlook.md)|
|[Copy](noteitem-copy-method-outlook.md)|
|[Delete](noteitem-delete-method-outlook.md)|
|[Display](noteitem-display-method-outlook.md)|
|[Move](noteitem-move-method-outlook.md)|
|[PrintOut](noteitem-printout-method-outlook.md)|
|[Save](noteitem-save-method-outlook.md)|
|[SaveAs](noteitem-saveas-method-outlook.md)|

## Properties



|**Name**|
|:-----|
|[Application](noteitem-application-property-outlook.md)|
|[AutoResolvedWinner](noteitem-autoresolvedwinner-property-outlook.md)|
|[Body](noteitem-body-property-outlook.md)|
|[Categories](noteitem-categories-property-outlook.md)|
|[Class](noteitem-class-property-outlook.md)|
|[Conflicts](noteitem-conflicts-property-outlook.md)|
|[CreationTime](noteitem-creationtime-property-outlook.md)|
|[DownloadState](noteitem-downloadstate-property-outlook.md)|
|[EntryID](noteitem-entryid-property-outlook.md)|
|[GetInspector](noteitem-getinspector-property-outlook.md)|
|[Height](noteitem-height-property-outlook.md)|
|[IsConflict](noteitem-isconflict-property-outlook.md)|
|[ItemProperties](noteitem-itemproperties-property-outlook.md)|
|[LastModificationTime](noteitem-lastmodificationtime-property-outlook.md)|
|[Left](noteitem-left-property-outlook.md)|
|[MarkForDownload](noteitem-markfordownload-property-outlook.md)|
|[MessageClass](noteitem-messageclass-property-outlook.md)|
|[Parent](noteitem-parent-property-outlook.md)|
|[PropertyAccessor](noteitem-propertyaccessor-property-outlook.md)|
|[Saved](noteitem-saved-property-outlook.md)|
|[Session](noteitem-session-property-outlook.md)|
|[Size](noteitem-size-property-outlook.md)|
|[Subject](noteitem-subject-property-outlook.md)|
|[Top](noteitem-top-property-outlook.md)|
|[Width](noteitem-width-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
