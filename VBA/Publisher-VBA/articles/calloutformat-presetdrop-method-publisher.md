---
title: "Метод CalloutFormat.PresetDrop (издатель)"
keywords: vbapb10.chm2490387
f1_keywords: vbapb10.chm2490387
ms.prod: publisher
api_name: Publisher.CalloutFormat.PresetDrop
ms.assetid: a709e54a-d08a-f83c-a0bf-bcdcfe6434cd
ms.date: 06/08/2017
ms.openlocfilehash: 3eee7a2c6c3a45c748691c669e516fc9120d16e5
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="calloutformatpresetdrop-method-publisher"></a>Метод CalloutFormat.PresetDrop (издатель)

Указывает ли линии выноски присоединяется к верхней, нижней или центр выноски текстового поля или его подключает на момент, который является определенное расстояние между верхней или нижней части текстового поля.


## <a name="syntax"></a>Синтаксис

 _выражение_. **PresetDrop** ( **_DropType_**)

 переменная _expression_A, представляет собой объект- **CalloutFormat** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|DropType|Обязательное свойство.| **MsoCalloutDropType**|Начальная позиция линии выноски по отношению к тексту, ограничивающий прямоугольник.|

## <a name="remarks"></a>Заметки

Параметр DropType может быть одной из констант **MsoCalloutDropType** объявлена в библиотеке типов, Microsoft Office и показаны в следующей таблице.



| **msoCalloutDropBottom**|| **msoCalloutDropCenter**|| **msoCalloutDropCustom**|| **msoCalloutDropTop**|

## <a name="example"></a>Пример

В этом примере указывается, что линии выноски с подключением в начало текста, ограничивающий прямоугольник для первой фигуры в активной публикации. Для обеспечения работы примера фигуры должен быть выноске.


```vb
ActiveDocument.Pages(1).Shapes(1).Callout.PresetDrop DropType:=msoCalloutDropTop
```

В этом примере для переключения между два предварительно падения для первой фигуры одно в активной публикации. Для обеспечения работы примера фигуры должен быть выноске.




```vb
With ActiveDocument.Pages(1).Shapes(1).Callout 
 Select Case .DropType 
 Case msoCalloutDropTop 
 .PresetDrop DropType:=msoCalloutDropBottom 
 Case msoCalloutDropBottom 
 .PresetDrop DropType:=msoCalloutDropTop 
 End Select 
End With 

```


