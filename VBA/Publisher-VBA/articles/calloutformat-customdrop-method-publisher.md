---
title: "Метод CalloutFormat.CustomDrop (издатель)"
keywords: vbapb10.chm2490385
f1_keywords: vbapb10.chm2490385
ms.prod: publisher
api_name: Publisher.CalloutFormat.CustomDrop
ms.assetid: 65fc7309-acd0-5bdd-6bb0-1b6c41968775
ms.date: 06/08/2017
ms.openlocfilehash: 8ce55fbd8583d8e2d7624fd88dbf327cab200d05
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="calloutformatcustomdrop-method-publisher"></a>Метод CalloutFormat.CustomDrop (издатель)

Задает расстояние по вертикали от края текста, ограничивающий прямоугольник в то место, где линии выноски подключает текстовое поле.


## <a name="syntax"></a>Синтаксис

 _выражение_. **CustomDrop** ( **_Удалить_**)

 переменная _expression_A, представляет собой объект- **CalloutFormat** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Поместите|Обязательное свойство.| **Variant**|Расстояние размещения сообщений. Числовые значения вычисляются в точках; строк может быть в любой устройств, поддерживаемых Microsoft Publisher (например, «2,5 дюйма»).|

## <a name="remarks"></a>Заметки

Расстояние размещения обычно отсчитывается от верхней части текстового поля. Тем не менее если свойство **[AutoAttach](calloutformat-autoattach-property-publisher.md)** имеет значение **True** , а поле — слева от происхождения линии выноски (месте, на который указывает выноски) раскрывающегося расстояния измеряется в нижней части текстового поля.


## <a name="example"></a>Пример

В этом примере задается расстояние перетаскивания 14 пунктов, а также определяет всегда измеряется раскрывающегося расстояния сверху. Для обеспечения работы примера третий фигуры в активной публикации должен быть выноске.


```vb
With ActiveDocument.Pages(1).Shapes(3).Callout 
 .CustomDrop Drop:=14 
 .AutoAttach = False 
End With 

```


