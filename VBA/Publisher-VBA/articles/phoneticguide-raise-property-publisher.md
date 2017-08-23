---
title: "Свойство PhoneticGuide.Raise (издатель)"
keywords: vbapb10.chm6160389
f1_keywords: vbapb10.chm6160389
ms.prod: publisher
api_name: Publisher.PhoneticGuide.Raise
ms.assetid: 8c7bd7e9-1b63-ded0-5021-99995296ad08
ms.date: 06/08/2017
ms.openlocfilehash: f239814ee840b43c2290cc3eef9754efdeaa6a4c
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="phoneticguideraise-property-publisher"></a>Свойство PhoneticGuide.Raise (издатель)

Возвращает значение **типа Variant** , указывающее расстояние между верхней части основного текста и в нижней части текста руководство по. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Вызывает**

 переменная _expression_A, представляет собой объект- **PhoneticGuide** .


### <a name="return-value"></a>Возвращаемое значение

Variant


## <a name="remarks"></a>Заметки

Числовой набор значений находятся в точках; строк может быть любой единицы измерения, поддерживаются в Microsoft Publisher. Возвращаемые значения всегда находятся в пунктах.


## <a name="example"></a>Пример

Следующий пример помещает фонетическое руководство для фигуры одно в пунктах пять активная публикация над основной текст.


```vb
Dim phoGuide As PhoneticGuide 
 
Set phoGuide = ActiveDocument.Pages(1).Shapes(1) _ 
 .TextFrame.TextRange.Fields(1).PhoneticGuide 
 
With phoGuide 
 .Raise = 5 
End With
```


