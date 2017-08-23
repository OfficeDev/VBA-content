---
title: "Метод Shape.ScaleWidth (издатель)"
keywords: vbapb10.chm2228262
f1_keywords: vbapb10.chm2228262
ms.prod: publisher
api_name: Publisher.Shape.ScaleWidth
ms.assetid: 07dcc04e-cb84-9c69-c589-87c0ff0bb147
ms.date: 06/08/2017
ms.openlocfilehash: a84faa4866e381e34077c63d3e0c8266bdcd0e68
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapescalewidth-method-publisher"></a>Метод Shape.ScaleWidth (издатель)

Масштабирование ширины фигуры с указанного коэффициента. Для изображений и объекты OLE можно указать, следует ли масштабировать фигуры относительно исходного размера или относительно текущего размера.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ScaleWidth** ( **_Коэффициент_** **_RelativeToOriginalSize_** **_fScale_**)

 переменная _expression_A, представляющий объект **фигуры** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Коэффициент|Обязательное свойство.| **Один**|Указывает отношение между ширина формы при изменении размера и текущей или исходной ширины. Например чтобы сделать прямоугольник более 50 процентов, укажите 1,5 для этого аргумента.|
|RelativeToOriginalSize|Обязательное свойство.| **MsoTristate**| Указывает, следует ли масштабировать относительно размер исходного или текущего объекта.|
|fScale|Необязательный| **MsoScaleFrom**|Часть фигуры, сохраняет его положение при масштабировании фигуры.|

## <a name="remarks"></a>Заметки

Параметр RelativeToOriginalSize может быть одной из констант **MsoTriState** объявлена в библиотеке типов, Microsoft Office и показаны в следующей таблице.



|**Константы**|**Описание**|
|:-----|:-----|
| **msoFalse**|Масштабируется фигуры относительно его текущий размер.|
| **msoTrue**|Масштабируется фигуры относительно исходного размера.|
Параметр fScale может быть одной из констант **MsoScaleFrom** объявлена в библиотеке типов, Microsoft Office и показаны в следующей таблице. Значение по умолчанию — **msoScaleFromTopLeft**.



| **msoScaleFromBottomRight**|| **msoScaleFromMiddle**|| **msoScaleFromTopLeft**| Фигуры, отличный от изображения и объекты OLE всегда масштабируются их текущей ширине; значение RelativeToOriginalSize **msoTrue** для фигур, отличный от изображения или объекты OLE приводит к ошибке.

Используйте метод **[ScaleHeight](shape-scaleheight-method-publisher.md)** масштабирование высоту фигуры.


## <a name="example"></a>Пример

В этом примере масштабирование все изображения и объекты OLE на первой странице active публикации до 175 процентов их исходной высоты и ширины и масштабов всех фигур 175% от их текущего высоту и ширину.


```vb
' Looping variable. 
Dim shpLoop As Shape 
 
' Loop through all the shapes on the first page. 
For Each shpLoop In ActiveDocument.Pages(1).Shapes 
 With shpLoop 
 Select Case .Type 
 ' If the shape is a picture or OLE object, 
 ' scale relative to original size. 
 Case pbPicture, pbLinkedPicture, _ 
 pbEmbeddedOLEObject, pbLinkedOLEObject, _ 
 pbOLEControlObject 
 .ScaleHeight Factor:=1.75, _ 
 RelativeToOriginalSize:=True 
 .ScaleWidth Factor:=1.75, _ 
 RelativeToOriginalSize:=True 
 ' If the shape is not a picture or OLE object, 
 ' scale relative to the current size. 
 Case Else 
 .ScaleHeight Factor:=1.75, _ 
 RelativeToOriginalSize:=False 
 .ScaleWidth Factor:=1.75, _ 
 RelativeToOriginalSize:=False 
 End Select 
 End With 
Next shpLoop 

```


