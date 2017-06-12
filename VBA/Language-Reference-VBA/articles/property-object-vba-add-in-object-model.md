---
title: Property Object (VBA Add-In Object Model)
keywords: vbob6.chm102045
f1_keywords:
- vbob6.chm102045
ms.prod: office
ms.assetid: 231018ff-4e74-fc67-a69b-0988e5b7517d
ms.date: 06/08/2017
---


# Property Object (VBA Add-In Object Model)



Represents the [properties](vbe-glossary.md) of an object that are visible in the[Properties window](vbe-glossary.md) for any given component.
 **Remarks**
Use  **Value** property of the **Property** object to return or set the value of a property of a component.
At a minimum, all components have a  **Name** property. Use the **Value** property of the **Property** object to return or set the value of a property. The **Value** property returns a[Variant](vbe-glossary.md) of the appropriate type. If the value returned is an object, the **Value** property returns the **Properties** collection that contains **Property** objects representing the individual properties of the object. You can access each of the **Property** objects by using the **Item** method on the returned **Properties** collection.
If the value returned by the  **Property** object is an object, you can use the **Object** property to set the **Property** object to a new object.

