---
title: "Метод ShapeRange.ScaleHeight (издатель)"
keywords: vbapb10.chm2293797
f1_keywords: vbapb10.chm2293797
ms.prod: publisher
api_name: Publisher.ShapeRange.ScaleHeight
ms.assetid: 54058fe5-d922-0ea9-08e8-99fff89bde55
ms.date: 06/08/2017
ms.openlocfilehash: 04e02ab50ef7c08d08d2ad8db019c94bed2d4507
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shaperangescaleheight-method-publisher"></a>Метод ShapeRange.ScaleHeight (издатель)

Масштабирование высота формы с указанным коэффициентом. Для изображений и объекты OLE можно указать, следует ли масштабировать фигуры относительно исходного размера или относительно текущего размера.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ScaleHeight** ( **_Коэффициент_** **_RelativeToOriginalSize_** **_fScale_**)

 переменная _expression_A, представляющий объект **ShapeRange** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Коэффициент|Обязательное свойство.| **Один**|Указывает отношение между высота формы при изменении размера и высота текущей или исходной. Например чтобы сделать прямоугольник более 50 процентов, укажите 1,5 для этого аргумента.|
|RelativeToOriginalSize|Обязательное свойство.| **MsoTriState**|Указывает, следует ли масштабировать относительно размер исходного или текущего объекта.|
|fScale|Необязательный| **MsoScaleFrom**|Часть фигуры, сохраняет его положение при масштабировании фигуры.|

## <a name="remarks"></a>Заметки

Параметр RelativeToOriginalSize может быть одной из констант **MsoTriState** объявлена в библиотеке типов, Microsoft Office и показаны в следующей таблице.



|**Константы**|**Описание**|
|:-----|:-----|
| **msoFalse**|Масштабируется фигуры относительно его текущий размер.|
| **msoTrue**|Масштабируется фигуры относительно исходного размера.|
Параметр fScale может быть одной из констант **MsoScaleFrom** объявлена в библиотеке типов, Microsoft Office и показаны в следующей таблице. Значение по умолчанию — **msoScaleFromTopLeft**.



| **msoScaleFromBottomRight**|| **msoScaleFromMiddle**|| **msoScaleFromTopLeft**| Фигуры, отличный от изображения и объекты OLE всегда масштабируются их текущей высоте; значение RelativeToOriginalSize **msoTrue** для фигур, отличный от изображения или объекты OLE приводит к ошибке.

Используйте метод **[ScaleWidth](shape-scalewidth-method-publisher.md)** масштабирование ширины фигуры.


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


