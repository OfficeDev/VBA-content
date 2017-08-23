---
title: "Метод Shapes.AddCallout (издатель)"
keywords: vbapb10.chm2162704
f1_keywords: vbapb10.chm2162704
ms.prod: publisher
api_name: Publisher.Shapes.AddCallout
ms.assetid: bbf5f913-fcf0-b700-0c7e-9f0bdc7c6aea
ms.date: 06/08/2017
ms.openlocfilehash: 94910d40aa383d911a66523049798b47ccb07122
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapesaddcallout-method-publisher"></a>Метод Shapes.AddCallout (издатель)

Добавление нового объекта **[Shape](shape-object-publisher.md)** , представляющее без границ выноску определенной коллекции **[фигур](shapes-object-publisher.md)** .


## <a name="syntax"></a>Синтаксис

 _выражение_. **AddCallout** ( **_Тип_**, **_слева_**, **_в начало_**, **_Width_**, **_Height_**)

 переменная _expression_A, представляет собой объект- **фигур** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Тип|Обязательный| **MsoCalloutType**|Тип линии выноски.|
|Слева|Обязательное свойство.| **Variant**|Положение левого края фигуры, представляющей выноску.|
|Вверх|Обязательное свойство.| **Variant**|Положение верхнего края фигуры, представляющей выноску.|
|Width|Обязательное свойство.| **Variant**|Ширина формы, представляющее выноску.|
|Height|Обязательное свойство.| **Variant**|Высота фигуры, представляющей выноску.|

### <a name="return-value"></a>Возвращаемое значение

Shape


## <a name="remarks"></a>Заметки

Аргументы слева, Top, ширину и высоту числовые значения вычисляются в точках; строк может быть в любой устройств, поддерживаемых Microsoft Publisher (например, «2,5 дюйма»).

Параметр типа может иметь одно из следующих констант **MsoCalloutType** .



| **msoCalloutOne**| Линия выноски сегмента одним горизонтальный или вертикальный. | | **msoCalloutTwo**| Свободно вращающимся линии выноски одним сегмент. | | **msoCalloutThree**| Линии выноски два сегмента. | | **msoCalloutFour**| Линии выноски три сегмента. |

## <a name="example"></a>Пример

Следующий пример добавляет новую строку свободно вращающимся выноски для первой страницы публикации active.


```vb
Dim shpCallout As Shape 
 
Set shpCallout = ActiveDocument.Pages(1).Shapes.AddCallout _ 
 (Type:=msoCalloutTwo, _ 
 Left:=144, Top:=216, _ 
 Width:=36, Height:=72)
```


