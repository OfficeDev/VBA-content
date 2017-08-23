---
title: "Метод ShapeRange.ZOrder (издатель)"
keywords: vbapb10.chm2293808
f1_keywords: vbapb10.chm2293808
ms.prod: publisher
api_name: Publisher.ShapeRange.ZOrder
ms.assetid: 2043f78c-ab83-e719-c3b5-5d75edcf1593
ms.date: 06/08/2017
ms.openlocfilehash: af4c148e9018bcbb4ffdd89f91e25d30481868d6
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shaperangezorder-method-publisher"></a>Метод ShapeRange.ZOrder (издатель)

Перемещает указанный фигуры на переднем плане или из-за других фигур в коллекции (то есть, изменяется фигуры позицию в z порядке).


## <a name="syntax"></a>Синтаксис

 _выражение_. **Метод ZOrder** ( **_ZOrderCmd_**)

 переменная _expression_A, представляющий объект **ShapeRange** .


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


