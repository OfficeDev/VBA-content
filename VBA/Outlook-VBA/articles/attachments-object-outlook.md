---
title: Attachments Object (Outlook)
keywords: vbaol11.chm169
f1_keywords:
- vbaol11.chm169
ms.prod: outlook
api_name:
- Outlook.Attachments
ms.assetid: 4cc96a5f-a822-8ad5-6f61-e996bee8ba22
ms.date: 06/08/2017
---


# Attachments Object (Outlook)

Contains a set of  **[Attachment](http://msdn.microsoft.com/library/3e11582b-ac90-0948-bc37-506570bb287b%28Office.15%29.aspx)** objects that represent the attachments in an Outlook item.


## Remarks

Use the  **[Attachments](http://msdn.microsoft.com/library/2843bef3-2ace-1cc0-1f15-c3fb776c3bf9%28Office.15%29.aspx)** property to return the **Attachments** collection for any Outlook item (except notes).

Use the  **[Add](http://msdn.microsoft.com/library/e11980fd-e1fc-a0c3-cdd0-0e598988d3c2%28Office.15%29.aspx)** method to add an attachment to an item.

To ensure consistent results, always save an item before adding or removing objects in the  **Attachments** collection of the item.


## Example

The following Visual Basic for Applications (VBA) example creates a new mail message, attaches a Q496.xls as an attachment (not a link), and gives the attachment a descriptive caption.


```
Set myItem = Application.CreateItem(olMailItem) 
 
myItem.Save 
 
Set myAttachments = myItem.Attachments 
 
myAttachments.Add "C:\My Documents\Q496.xls", _ 
 
 olByValue, 1, "4th Quarter 1996 Results Chart"
```


## Methods



|**Name**|
|:-----|
|[Add](http://msdn.microsoft.com/library/e11980fd-e1fc-a0c3-cdd0-0e598988d3c2%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/2843bef3-2ace-1cc0-1f15-c3fb776c3bf9%28Office.15%29.aspx)|
|[Remove](http://msdn.microsoft.com/library/be49c973-b64e-84d9-1bf6-73b27a7e84f0%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/4ca29aab-f2dd-3625-b964-d9582cbd7fdf%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/29f722c7-7117-0827-1531-fa45d2b4b6b5%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/f25a85a0-298e-457d-b2b6-7f7ec18c6921%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/c8f54089-14b8-b8e2-8483-53e76b12aaf4%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/af206370-3d50-84de-187d-019126958b61%28Office.15%29.aspx)|

## See also


#### Other resources


[Attach a File to a Mail Item](http://msdn.microsoft.com/library/1d94629b-e713-92cb-32de-c8910612e861%28Office.15%29.aspx)
[Attach an Outlook Contact Item to an Email Message](http://msdn.microsoft.com/library/ae5240ad-dc3e-4499-8fd0-d8c2d90aa9ba%28Office.15%29.aspx)
[Limit the Size of an Attachment to an Outlook Email Message](http://msdn.microsoft.com/library/9a240e17-f715-482c-9a8b-c6be1144e15a%28Office.15%29.aspx)
[Modify an Attachment of an Outlook Email Message](http://msdn.microsoft.com/library/f5dac09a-272b-49d6-bf1e-82c3981260ed%28Office.15%29.aspx)
[Attachments Object Members](http://msdn.microsoft.com/library/cfdc1209-1b17-9b6c-122c-c07122d3aae1%28Office.15%29.aspx)
[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
