---
title: Selection Object (Outlook)
keywords: vbaol11.chm80
f1_keywords:
- vbaol11.chm80
ms.prod: outlook
api_name:
- Outlook.Selection
ms.assetid: 0b06a3ce-0445-db8f-e6e8-bb7bd469c50f
ms.date: 06/08/2017
---


# Selection Object (Outlook)

Contains the set of Outlook items currently selected in an explorer.


## Remarks

Use the  **[Selection](http://msdn.microsoft.com/library/11002043-9dab-a5ad-b36e-52ddb04c1859%28Office.15%29.aspx)** property to return the **Selection** collection from the **[Explorer](explorer-object-outlook.md)** object.


## Example

The following example returns a  **Selection** object from an **Explorer** object.


```
Set mySelectedItems = myExplorer.Selection
```


## Methods



|**Name**|
|:-----|
|[GetSelection](http://msdn.microsoft.com/library/c6af6665-d97d-3833-1014-5b43282bafc2%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/981b107a-14d7-2dd3-6449-2737b2801c3c%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/06ce9b99-1323-2611-dd3a-5646bb1b0ec8%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/a05de32a-2a2a-3579-bc47-545efaf92a8d%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/ea7a19d2-6261-ce07-97f3-ebe95489a265%28Office.15%29.aspx)|
|[Location](http://msdn.microsoft.com/library/8a2db72a-8db0-840e-349e-5d9d22f3affb%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/a081601f-a0ee-d998-f0e9-0193f9db843e%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/22390a36-a51c-615d-a646-45e5aa7d253f%28Office.15%29.aspx)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
[Selection Object Members](http://msdn.microsoft.com/library/c79922d4-aa76-ff48-f163-8161fa1ae0a8%28Office.15%29.aspx)
