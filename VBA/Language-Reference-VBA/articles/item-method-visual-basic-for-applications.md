---
title: Item Method (Visual Basic for Applications)
keywords: vblr6.chm1014019
f1_keywords:
- vblr6.chm1014019
ms.prod: office
ms.assetid: 6850a534-f6cc-e4be-3fc9-4975d1cff775
ms.date: 06/08/2017
---


# Item Method (Visual Basic for Applications)



Returns a specific [member](vbe-glossary.md) of a **Collection** object either by position or by key.
 **Syntax**
 _object_**.Item(**_index_**)**
The  **Item** method syntax has the following object qualifier and part:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. An [object expression](vbe-glossary.md) that evaluates to an object in the Applies To list.|
| _index_|Required. An [expression](vbe-glossary.md) that specifies the position of a member of the[collection](vbe-glossary.md). If a [numeric expression](vbe-glossary.md),  _index_ must be a number from 1 to the value of the collection's **Count** property. If a[string expression](vbe-glossary.md),  _index_ must correspond to the **_key_**[argument](vbe-glossary.md) specified when the member referred to was added to the collection.|
 **Remarks**
If the value provided as  _index_ doesn't match any existing member of the collection, an error occurs.
The  **Item** method is the default method for a collection. Therefore, the following lines of code are equivalent:



```
Print MyCollection(1)
Print MyCollection.Item(1)

```


## Example

This example uses the  **Item** method to retrieve a reference to an object in a collection. Assuming `Birthdays` is a **Collection** object, the following code retrieves from the collection references to the objects representing Bill Smith's birthday and Adam Smith's birthday, using the keys "SmithBill" and "SmithAdam" as the _index_ arguments. Note that the first call explicitly specifies the **Item** method, but the second does not. Both calls work because the **Item** method is the default for a **Collection** object. The references, assigned to `SmithBillBD` and `SmithAdamBD` using **Set**, can be used to access the properties and methods of the specified objects. To run this code, create the collection and populate it with at least the two referenced members.


```vb
Dim SmithBillBD As Object
Dim SmithAdamBD As Object
Dim Birthdays
Set SmithBillBD = Birthdays.Item("SmithBill")
Set SmithAdamBD = Birthdays("SmithAdam")
```


