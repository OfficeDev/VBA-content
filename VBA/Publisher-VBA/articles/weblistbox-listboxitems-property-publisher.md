---
title: "Свойство WebListBox.ListBoxItems (издатель)"
keywords: vbapb10.chm4063235
f1_keywords: vbapb10.chm4063235
ms.prod: publisher
api_name: Publisher.WebListBox.ListBoxItems
ms.assetid: 642a4592-35af-99fa-ee96-6bd8517c618f
ms.date: 06/08/2017
ms.openlocfilehash: 349f0b767fb6e2a9c78543da9a435db6ce6d3b4c
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="weblistboxlistboxitems-property-publisher"></a>Свойство WebListBox.ListBoxItems (издатель)

Возвращает объект **[WebListBoxItems](weblistboxitems-object-publisher.md)** , представляющий элементов в элемент управления списка Web.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ListBoxItems**

 переменная _expression_A, представляет собой объект- **WebListBox** .


### <a name="return-value"></a>Возвращаемое значение

WebListBoxItems


## <a name="example"></a>Пример

В этом примере создается новый элемент управления полем Web списка и добавляет пять новых элементов списка.


```vb
Sub NewListBoxItems() 
 Dim intCount As Integer 
 With ActiveDocument.Pages(1).Shapes.AddWebControl _ 
 (Type:=pbWebControlListBox, Left:=100, _ 
 Top:=100, Width:=150, Height:=100).WebListBox 
 .MultiSelect = msoTrue 
 With .ListBoxItems 
 For intCount = 1 To .Count 
 .Delete (1) 
 Next 
 .AddItem Item:="Yellow" 
 .AddItem Item:="Red" 
 .AddItem Item:="Blue" 
 .AddItem Item:="Green" 
 .AddItem Item:="Black" 
 End With 
 End With 
End Sub
```


