---
title: UserProperties Object (Outlook)
keywords: vbaol11.chm202
f1_keywords:
- vbaol11.chm202
ms.prod: outlook
api_name:
- Outlook.UserProperties
ms.assetid: 20b49c86-d74f-9bda-382c-559af278c148
ms.date: 06/08/2017
---


# UserProperties Object (Outlook)

Contains  **[UserProperty](userproperty-object-outlook.md)** objects that represent the custom properties of an Outlook item.


## Remarks

Use the  **UserProperties** property to return the **UserProperties** object for an Outlook item. This applies to all Outlook items except for the **[NoteItem](http://msdn.microsoft.com/library/ddf5baaa-6e13-a6fb-96e8-311e7761fa98%28Office.15%29.aspx)**.

Use the  **[Add](http://msdn.microsoft.com/library/88b86622-2234-77be-41e7-b76b0b3a75ad%28Office.15%29.aspx)** method to create a new **UserProperty** for an item and add it to the **UserProperties** object. The **Add** method allows you to specify a name and type for the new property. When you create a new property, it can also be added as a custom field to the folder that contains the item (using the same name as the property) by setting the _AddToFolderFields_ parameter to **True** when calling the **Add** method. That field can then be used as a column in folder views.

Use  **UserProperties** ( _index_ ), where _index_ is a name or one-based index number, to return a single **[UserProperty](userproperty-object-outlook.md)** object.

You can use the  **[UserDefinedProperties](http://msdn.microsoft.com/library/4293bcb8-855e-4c6d-9718-ba8c5862b3bd%28Office.15%29.aspx)** property of the **[Folder](folder-object-outlook.md)** object to retrieve and examine the definitions of custom item-level properties that a folder can display in a view.

To get or set multiple custom properties, use the  **[PropertyAccessor](propertyaccessor-object-outlook.md)** object instead of the **UserProperties** object for better performance.


## Example

The following example adds a custom text property named MyPropName to myItem.


```
Set myProp = myItem.UserProperties.Add("MyPropName", olText)
```


## Methods



|**Name**|
|:-----|
|[Add](http://msdn.microsoft.com/library/88b86622-2234-77be-41e7-b76b0b3a75ad%28Office.15%29.aspx)|
|[Find](http://msdn.microsoft.com/library/3b71ce5a-4bb0-fdab-a24e-02c631816b80%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/3e024200-0014-6a7d-dd34-9fcd0d2dd292%28Office.15%29.aspx)|
|[Remove](http://msdn.microsoft.com/library/47b77e76-3164-12d1-bf08-fa11847eafcb%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/cf34337d-7087-7a71-e13b-9f97beb605ca%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/be47a8e7-a5cb-2b9b-6fec-2e1090329f6b%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/6e316d8a-68b5-f25a-c3d2-4d72a054b027%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/35c01dd0-bec0-ece8-59fd-80daf1989e98%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/0cd76318-80c6-4cfc-3aca-32e385ff6b88%28Office.15%29.aspx)|

## See also


#### Other resources


[UserProperties Object Members](http://msdn.microsoft.com/library/b71f8a0b-3951-cfb0-89f2-df8851f3993d%28Office.15%29.aspx)
[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
