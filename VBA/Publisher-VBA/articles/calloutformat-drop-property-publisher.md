---
title: "Свойство CalloutFormat.Drop (издатель)"
keywords: vbapb10.chm2490629
f1_keywords: vbapb10.chm2490629
ms.prod: publisher
api_name: Publisher.CalloutFormat.Drop
ms.assetid: 7878a6a6-9c7c-dfd0-ef1b-d56a5aab6a18
ms.date: 06/08/2017
ms.openlocfilehash: 7a0730c8bbbcdd54c6ace23d19510b71dbf8cb55
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="calloutformatdrop-property-publisher"></a>Свойство CalloutFormat.Drop (издатель)

Для выносок с явным образом установлен перетащите значение, данное свойство возвращает расстояние по вертикали от края текста, ограничивающий прямоугольник в то место, где линии выноски подключает текстовое поле. Это расстояние отсчитывается от верхней части текстового поля, если не **AutoAttach** задано значение **True,** а поле слева от происхождения линии выноски (где выноски указывает). В этом случае раскрывающегося расстояния измеряется в нижней части текстового поля. Только для чтения **Variant**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Поместите**

 переменная _expression_A, представляет собой объект- **CalloutFormat** .


### <a name="return-value"></a>Возвращаемое значение

Variant


## <a name="remarks"></a>Заметки

Числовые значения вычисляются в точках; строк может быть в любой устройств, поддерживаемых Microsoft Publisher (например, «2,5 дюйма»).

Используйте метод **[CustomDrop](calloutformat-customdrop-method-publisher.md)** для задания значения этого свойства.

Значение этого свойства точно отражает положение вложения выноски строки в текстовом поле только в том случае, если выноски имеет явным образом установлен перетащите значение, то есть, если значение свойства **[DropType](calloutformat-droptype-property-publisher.md)** **msoCalloutDropCustom**.


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


