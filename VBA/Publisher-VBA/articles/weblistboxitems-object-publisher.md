---
title: "Объект WebListBoxItems (издатель)"
keywords: vbapb10.chm4194303
f1_keywords: vbapb10.chm4194303
ms.prod: publisher
api_name: Publisher.WebListBoxItems
ms.assetid: 6d1b6755-426b-b518-c95c-7b30f9acceba
ms.date: 06/08/2017
ms.openlocfilehash: 5aca9cc5d5d37539e8bd732671ed4abb398ba166
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="weblistboxitems-object-publisher"></a>Объект WebListBoxItems (издатель)

Представляет элементы в элемент управления списка Web.
 


## <a name="example"></a>Пример

Используйте свойство **[ListBoxItems](weblistbox-listboxitems-property-publisher.md)** для доступа к элементам в поле со списком Web. Использование метода **[AddItem](weblistboxitems-additem-method-publisher.md)** коллекции **WebListBoxItems** для добавления элементов в поле со списком Web. В этом примере создается новый список Web и добавляет несколько элементов. Обратите внимание, что при создании, элемент управления списка Web содержит три элемента по умолчанию. В этом примере включает в себя процедуры, в котором удаляются поля элементов списка по умолчанию, прежде чем добавлять новые элементы.
 

 

```
Sub CreateWebListBox() 
 Dim intCount As Integer 
 With ActiveDocument.Pages(1).Shapes 
 With .AddWebControl(Type:=pbWebControlListBox, Left:=100, _ 
 Top:=150, Width:=300, Height:=72).WebListBox 
 .MultiSelect = msoFalse 
 With .ListBoxItems 
 For intCount = 1 To .Count 
 .Delete (1) 
 Next 
 .AddItem Item:="Green" 
 .AddItem Item:="Purple" 
 .AddItem Item:="Red" 
 .AddItem Item:="Black" 
 End With 
 End With 
 End With 
End Sub
```


