---
title: "Свойство CalloutFormat.DropType (издатель)"
keywords: vbapb10.chm2490630
f1_keywords: vbapb10.chm2490630
ms.prod: publisher
api_name: Publisher.CalloutFormat.DropType
ms.assetid: fd4ec192-0732-e860-4ff8-e305aa0d90a9
ms.date: 06/08/2017
ms.openlocfilehash: 9286c1a397606df0355b535f6e99cbaca826255b
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="calloutformatdroptype-property-publisher"></a>Свойство CalloutFormat.DropType (издатель)

Возвращает константу **MsoCalloutDropType** , указывающее, где линии выноски подключает текстовое поле выноски. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **DropType**

 переменная _expression_A, представляет собой объект- **CalloutFormat** .


### <a name="return-value"></a>Возвращаемое значение

MsoCalloutDropType


## <a name="remarks"></a>Заметки

Значение свойства **DropType** может иметь одно из ** [MsoCalloutDropType](http://msdn.microsoft.com/library/0923e0a7-beb6-224f-6a87-85111f58ae3b%28Office.15%29.aspx)** объявленные константы в библиотеке типов, Microsoft Office.

Если удалить выноски тип — **msoCalloutDropCustom**, значения свойств **[Удалить](calloutformat-drop-property-publisher.md)** и **[AutoAttach](calloutformat-autoattach-property-publisher.md)** и относительное расположение текстовое поле выноски и выноски строка происхождение (где выноски указывает) используются для определения, где линии выноски подключает текстовое поле.

Используйте метод **[PresetDrop](calloutformat-presetdrop-method-publisher.md)** для задания значения этого свойства.


## <a name="example"></a>Пример

В этом примере заменяет перетаскивания для первой фигуры в активной публикации с одним из двух предварительно падения, в зависимости от того, является ли значение настраиваемого перетаскивания больше или меньше, чем половину высоты текстовое поле выноски. Для обеспечения работы примера фигуры должен быть выноске.


```vb
With ActiveDocument.Pages(1).Shapes(1).Callout 
 If .DropType = msoCalloutDropCustom Then 
 If .Drop < .Parent.Height / 2 Then 
 .PresetDrop DropType:=msoCalloutDropTop 
 Else 
 .PresetDrop DropType:=msoCalloutDropBottom 
 End If 
 End If 
End With 

```


