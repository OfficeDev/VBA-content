---
title: "Метод InlineShapes.Item (издатель)"
keywords: vbapb10.chm5767168
f1_keywords: vbapb10.chm5767168
ms.prod: publisher
api_name: Publisher.InlineShapes.Item
ms.assetid: 7cc4bb2a-e7d8-68c1-7d09-9b81a9d6b87a
ms.date: 06/08/2017
ms.openlocfilehash: 966cc85148674421178eaf8aecacc4c5dc47c350
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="inlineshapesitem-method-publisher"></a>Метод InlineShapes.Item (издатель)

Возвращает объект **[фигуры](shape-object-publisher.md)** , представляющий встроенная фигура, содержащихся в диапазон текста. Этот метод является элементом по умолчанию коллекции **InlineShapes** .


## <a name="syntax"></a>Синтаксис

 _выражение_. **Элемент** ( **_Индекс_**)

 переменная _expression_A, представляющий объект **InlineShapes** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|var|Обязательное свойство.| **Variant**|Индекс или имя возвращаемого объекта. Если **аргумент Index** имеет целое число, индекс в коллекции, основанный на 1. Если **аргумент Index** имеет строку, имя фигуры используется в качестве индекса. Если индекс или имя не представляет фигуры в коллекции, возвращается ошибка автоматизации.|

### <a name="return-value"></a>Возвращаемое значение

Shape


## <a name="example"></a>Пример

В этом примере выполняется поиск первую фигуру встроенного в диапазон текста и зеркальное отражение по вертикали.


```vb
Dim theShape As Shape 
 
Set theShape = ActiveDocument.Pages(1).Shapes(1) 
 
With theShape.TextFrame.Story.TextRange 
 With .InlineShapes.Item(1) 
 .Flip (msoFlipVertical) 
 End With 
End With
```


