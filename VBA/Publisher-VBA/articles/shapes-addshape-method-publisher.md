---
title: "Метод Shapes.AddShape (издатель)"
keywords: vbapb10.chm2162712
f1_keywords: vbapb10.chm2162712
ms.prod: publisher
api_name: Publisher.Shapes.AddShape
ms.assetid: 500d8cb3-f066-fdb6-09ac-b03c7822e8bd
ms.date: 06/08/2017
ms.openlocfilehash: 42acd0b3253a9b36d647ae072100e3c574127fee
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapesaddshape-method-publisher"></a>Метод Shapes.AddShape (издатель)

Добавляет новый объект **фигуры** , представляющее автофигуры к определенной коллекции **фигур** .


## <a name="syntax"></a>Синтаксис

 _выражение_. **AddShape** ( **_Тип_**, **_слева_**, **_в начало_**, **_Width_**, **_Height_**)

 переменная _expression_A, представляет собой объект- **фигур** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Тип|Обязательный| **MsoAutoShapeType**|Тип автофигуры иными способами. Полный список MsoAutoShapeType константы в разделе обозреватель объектов.|
|Слева|Обязательное свойство.| **Variant**|Положение левого края фигуры, представляющей автофигуры.|
|Вверх|Обязательное свойство.| **Variant**|Положение верхнего края фигуры, представляющей автофигуры.|
|Width|Обязательное свойство.| **Variant**|Ширина формы, представляющее автофигуры.|
|Height|Обязательное свойство.| **Variant**|Высота фигуры, представляющей автофигуры.|

### <a name="return-value"></a>Возвращаемое значение

Shape


## <a name="remarks"></a>Заметки

Для **_слева_**, **_Top_**, **_ширину_**и **_высоту_** аргументы числовые значения вычисляются в точках; строк может быть в любой устройств, поддерживаемых Microsoft Publisher (например, «2,5 дюйма»).


## <a name="example"></a>Пример

Следующий пример добавляет прямоугольник для первой страницы active публикации.


```vb
Dim shpShape As Shape 
 
Set shpShape = ActiveDocument.Pages(1).Shapes.AddShape _ 
 (Type:=msoShapeRectangle, _ 
 Left:=144, Top:=144, _ 
 Width:=72, Height:=144) 

```


