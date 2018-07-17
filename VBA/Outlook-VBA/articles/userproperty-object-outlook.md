---
title: UserProperty Object (Outlook)
keywords: vbaol11.chm212
f1_keywords:
- vbaol11.chm212
ms.prod: outlook
api_name:
- Outlook.UserProperty
ms.assetid: c94f642f-4368-d775-a79f-ce6c39bfe1fd
ms.date: 06/08/2017
---


# UserProperty Object (Outlook)

Represents a custom property of an Outlook item.


## Remarks

Use  **[UserProperties](http://msdn.microsoft.com/library/702ae502-d427-eeaf-ddd0-ff9749e7148c%28Office.15%29.aspx)** ( _index_ ), where _index_ is a name or index number, to return a single **UserProperty** object.

Use the  **[Add](http://msdn.microsoft.com/library/88b86622-2234-77be-41e7-b76b0b3a75ad%28Office.15%29.aspx)** method to create a new **UserProperty** for an item and add it to the **[UserProperties](userproperties-object-outlook.md)** object. The **Add** method allows you to specify a name and type for the new property.




 **Note**  When you create a custom property, a field is added in the folder that contains the item (using the same name as the property). That field can be used as a column in folder views.


## Example

The following example adds a custom text property named MyPropName.


```
Set myProp = myItem.UserProperties.Add("MyPropName", olText)
```


## Methods



|**Name**|
|:-----|
|[Delete](http://msdn.microsoft.com/library/6b1da165-f3d9-0a44-4582-3b468896a911%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/8796ad9a-dc97-72b4-9bcf-14cb9196335a%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/06f17b5f-0d42-6f7e-637c-5754a74aea9c%28Office.15%29.aspx)|
|[Formula](http://msdn.microsoft.com/library/91d2a104-8a93-a1e3-f31a-a0351153496d%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/7587062a-9cac-ed81-90a6-f1f0f089e757%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/8d584074-d3b0-ecbd-430e-afa083369773%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/181d0aad-9b03-9cce-b6dd-33a290d57ee9%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/d1eea53e-c46d-8dad-94cd-9338091b4ffd%28Office.15%29.aspx)|
|[ValidationFormula](http://msdn.microsoft.com/library/1420a7d9-2d10-ea1a-a893-e573f93919ad%28Office.15%29.aspx)|
|[ValidationText](http://msdn.microsoft.com/library/f2defd65-2c48-a24a-8cdc-a05b752cde53%28Office.15%29.aspx)|
|[Value](http://msdn.microsoft.com/library/9f313262-ffd4-3245-f516-bc2d62d6f33a%28Office.15%29.aspx)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
[UserProperty Object Members](http://msdn.microsoft.com/library/5c57c335-62b1-8d66-b93c-c56be823a85e%28Office.15%29.aspx)
