---
title: Pages Object (Outlook)
keywords: vbaol11.chm390
f1_keywords:
- vbaol11.chm390
ms.prod: outlook
api_name:
- Outlook.Pages
ms.assetid: ed4dd77e-b339-7f43-d036-c02daa69d5b8
ms.date: 06/08/2017
---


# Pages Object (Outlook)

Contains pages that represent the pages of an Inspector window.


## Remarks

Every  **[Inspector](inspector-object-outlook.md)** object has a **Pages** object defined, which is empty (count 0) if the Outlook item has never been customized before.

Use the  **[ModifiedFormPages](inspector-modifiedformpages-property-outlook.md)** property to return the **Pages** object from an **Inspector** object.

Use the  **[Add](pages-add-method-outlook.md)** method to create a custom page (you can add as many as 5 customizable pages). Use the ** _Name_** argument of the **Add** method to set the display name of the returned page. In addition to adding custom pages, you can use the _Name_ argument to return the main page of an **Inspector** object for modification.

Use  **ModifiedFormPages** ( _index_ ), where _index_ is the name or index number, to return a single page from a **Pages** object.


## Example



The following example returns the  **Pages** object for the active **Inspector**.




```
Set myPages = myItem.GetInspector.ModifiedFormPages
```

The following example returns a custom page with a default name (such as "Custom1").




```
Set myPage = myPages.Add
```

The following example returns a custom page named "My Page."






```
Set myPage = myPages.Add("My Page")
```

The following example returns the Message page if the Inspector contains a mail message.




```
Set myPage = myPages.Add("Message")
```

The following example returns the General (main) page if the inspector contains a contact.




```
Set myPage = myPages.Add("General")
```


## Methods



|**Name**|
|:-----|
|[Add](pages-add-method-outlook.md)|
|[Item](pages-item-method-outlook.md)|
|[Remove](pages-remove-method-outlook.md)|

## Properties



|**Name**|
|:-----|
|[Application](pages-application-property-outlook.md)|
|[Class](pages-class-property-outlook.md)|
|[Count](pages-count-property-outlook.md)|
|[Parent](pages-parent-property-outlook.md)|
|[Session](pages-session-property-outlook.md)|

## See also


#### Other resources


[Object model (Outlook VBA reference)](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
