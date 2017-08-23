---
title: "Метод Shape.ZOrder (издатель)"
keywords: vbapb10.chm2228272
f1_keywords: vbapb10.chm2228272
ms.prod: publisher
api_name: Publisher.Shape.ZOrder
ms.assetid: 05143a2b-924e-b5a3-390d-9493627bfa9f
ms.date: 06/08/2017
ms.openlocfilehash: 2a5c84a52b4d2be455ed38395a8e55d3ea870d88
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapezorder-method-publisher"></a>Метод Shape.ZOrder (издатель)

Перемещает указанный фигуры на переднем плане или из-за других фигур в коллекции (то есть, изменяется фигуры позицию в z порядке).


## <a name="syntax"></a>Синтаксис

 _выражение_. **Метод ZOrder** ( **_ZOrderCmd_**)

 переменная _expression_A, представляющий объект **фигуры** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|ZOrderCmd|Обязательное свойство.| **MsoZOrderCmd**|Указывает, где для перемещения указанного фигуры относительно других фигур.|

### <a name="return-value"></a>Возвращаемое значение

Значение Nothing


## <a name="remarks"></a>Заметки

Параметр ZOrderCmd может быть одной из констант **MsoZOrderCmd** объявлена в библиотеке типов, Microsoft Office и показаны в следующей таблице.



| **msoBringForward**|| **msoBringInFrontOfText**|| **msoBringToFront**|| **msoSendBackward**|| **msoSendBehindText**|| **msoSendToBack**| Свойство [ZOrderPosition](shape-zorderposition-property-publisher.md)определяет фигуры текущую позицию в z порядке.


## <a name="example"></a>Пример

В этом примере добавляет овала active публикации и помещает Овал второй с обратной в z порядке при наличии по крайней мере один фигуры на странице.


```vb
With ActiveDocument.Pages(1).Shapes _ 
 .AddShape(Type:=msoShapeOval, _ 
 Left:=100, Top:=100, Width:=100, Height:=300) 
 While .ZOrderPosition > 2 
 .ZOrder ZOrderCmd:=msoSendBackward 
 Wend 
End With 

```


