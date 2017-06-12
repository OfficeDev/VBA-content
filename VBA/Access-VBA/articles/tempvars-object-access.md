---
title: TempVars Object (Access)
keywords: vbaac10.chm14073
f1_keywords:
- vbaac10.chm14073
ms.prod: access
api_name:
- Access.TempVars
ms.assetid: aa81b18b-5e9f-ae44-cbcf-55cf6e37b7f6
ms.date: 06/08/2017
---


# TempVars Object (Access)

Represents the collection of  **[TempVar](tempvar-object-access.md)** objects.


## Remarks

Use the  **[Add](http://msdn.microsoft.com/library/836e449c-35ff-4089-857a-403c9fc97592%28Office.15%29.aspx)** method or the[SetTempVar](http://msdn.microsoft.com/library/9c3b7bee-02c5-efbf-1276-4c4a1f7802d9%28Office.15%29.aspx) macro action to create a **TempVar** object.

Use the  **[Remove](http://msdn.microsoft.com/library/a9ab9ff2-5bfc-d001-f5eb-9929907bc1b2%28Office.15%29.aspx)** method or the[RemoveTempVar](http://msdn.microsoft.com/library/409fd836-4a53-cefd-4264-8cee0fa8ac52%28Office.15%29.aspx) macro action to delete a **TempVar** object from the **TempVars** collection.

Use the  **[RemoveAll](http://msdn.microsoft.com/library/1b278bda-9f28-8fd7-0408-3a2a4d3e1a74%28Office.15%29.aspx)** method or[RemoveAllTempVars](http://msdn.microsoft.com/library/409fd836-4a53-cefd-4264-8cee0fa8ac52%28Office.15%29.aspx) macro action to delete all **TempVar** objects from the **TempVars** collection.

The  **TempVars** collection can store up to 255 **TempVar** objects. If you do not remove a **TempVar** object, it will remain in memory until you close the database. It is a good practice to remove **TempVar** object variables when you are finished using them.

To refer to a  **TempVar** object in a collection by its ordinal number or by its **Name** property setting, use the following syntax form:


-  **TempVar** ![name]
    

## Methods



|**Name**|
|:-----|
|[Add](http://msdn.microsoft.com/library/836e449c-35ff-4089-857a-403c9fc97592%28Office.15%29.aspx)|
|[Remove](http://msdn.microsoft.com/library/a9ab9ff2-5bfc-d001-f5eb-9929907bc1b2%28Office.15%29.aspx)|
|[RemoveAll](http://msdn.microsoft.com/library/1b278bda-9f28-8fd7-0408-3a2a4d3e1a74%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/250a64f6-d0a2-d816-1211-c56d90de0e70%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/3d4bfc9c-3a7c-5470-0e11-8e88bb5014e6%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/b2b71b6c-cfb4-0b1d-2417-a71725584642%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/0dfb9feb-54ef-e15d-2569-1261f2ae3358%28Office.15%29.aspx)|

## See also


#### Other resources


[Access Object Model Reference](http://msdn.microsoft.com/library/2de134a4-6c5c-d2a3-8377-f4dd973ba650%28Office.15%29.aspx)
[TempVars Object Members](http://msdn.microsoft.com/library/5c83c870-c66c-8fd9-0ac6-06766b14a6fc%28Office.15%29.aspx)
