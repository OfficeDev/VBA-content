---
title: "Метод WebListBoxItems.AddItem (издатель)"
keywords: vbapb10.chm4128772
f1_keywords: vbapb10.chm4128772
ms.prod: publisher
api_name: Publisher.WebListBoxItems.AddItem
ms.assetid: 1c3af4d1-ed0b-60c6-b607-17712612cec2
ms.date: 06/08/2017
ms.openlocfilehash: 5b1d7af4747241e8b351bdf75c10e8f97edf0b7b
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="weblistboxitemsadditem-method-publisher"></a>Метод WebListBoxItems.AddItem (издатель)

Добавляет элемент управления списка Web элементов списка.


## <a name="syntax"></a>Синтаксис

 _выражение_. **AddItem** ( **_Элемент_**, **_индекса_**, **_SelectState_**, **_ItemValue_**)

 переменная _expression_A, представляет собой объект- **WebListBoxItems** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Item|Обязательное свойство.| **String**|Имя элемента, как оно отображается в списке.|
|Индекс|Необязательный| **Длинный**|Номер элемента списка. Если не указан индекс или находится вне диапазона индексов существующих элементов списка, новый элемент добавляется в конец списка. В противном случае — новый элемент вставляется в позиции, заданной параметром индекса и позицию индекса все элементы после увеличивается на один.|
|SelectState|Необязательный| **Boolean**| **Значение true,** Если элемент выбран, когда сначала отображается поле со списком. Значение по умолчанию — **False**.|
|ItemValue|Необязательный| **String**|Значение элемента списка. Если не указан, значение нового элемента будет совпадать с именем элемента.|

## <a name="remarks"></a>Заметки

При создании нового списка Web программными средствами содержит три элемента. Используйте метод **[Delete](weblistboxitems-delete-method-publisher.md)** , чтобы удалить их из списка.


## <a name="example"></a>Пример

В этом примере создается новый элемент управления полем списка в активной публикации, удаляет трех стандартных элементов списка и добавляет несколько элементов.


```vb
Sub AddListBoxItems() 
 Dim intCount As Integer 
 With ActiveDocument.Pages(1).Shapes.AddWebControl _ 
 (Type:=pbWebControlListBox, Left:=100, _ 
 Top:=100, Width:=150, Height:=100) 
 With .WebListBox.ListBoxItems 
 For intCount = 1 To .Count 
 .Delete (1) 
 Next 
 .AddItem Item:="Green" 
 .AddItem Item:="Yellow" 
 .AddItem Item:="Red" 
 .AddItem Item:="Blue" 
 .AddItem Item:="Purple" 
 .AddItem Item:="Chartreuse" 
 .AddItem Item:="Pink" 
 .AddItem Item:="Olive" 
 End With 
 End With 
End Sub
```


