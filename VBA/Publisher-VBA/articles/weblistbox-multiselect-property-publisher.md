---
title: "Свойство WebListBox.MultiSelect (издатель)"
keywords: vbapb10.chm4063236
f1_keywords: vbapb10.chm4063236
ms.prod: publisher
api_name: Publisher.WebListBox.MultiSelect
ms.assetid: cc81682f-5212-0912-d979-16567c2dc42b
ms.date: 06/08/2017
ms.openlocfilehash: 9fb1da4483d6e9320bacecd00eb3cc2b0446a489
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="weblistboxmultiselect-property-publisher"></a>Свойство WebListBox.MultiSelect (издатель)

Указывает, может ли пользователь выбирать более одного элемента в элемент управления списка Web. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **MultiSelect**

 переменная _expression_A, представляет собой объект- **WebListBox** .


### <a name="return-value"></a>Возвращаемое значение

MsoTriState


## <a name="remarks"></a>Заметки

Значение свойства **MultiSelect** может иметь одно из **MsoTriState** константы объявляются в библиотеке типов Microsoft Office и показаны в следующей таблице.



|**Константы**|**Описание**|
|:-----|:-----|
| **msoFalse**| Указывает, что пользователь может выбрать только один элемент в элемент управления списка Web.|
| **msoTrue**| Указывает, что пользователь может выбрать более одного элемента в элемент управления списка Web.|

## <a name="example"></a>Пример

В этом примере добавьте элемент управления списка Web active публикации, добавление элементов к нему и указывает, что пользователь может выбрать несколько элементов.


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


