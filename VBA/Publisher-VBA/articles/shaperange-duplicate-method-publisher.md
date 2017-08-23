---
title: "Метод ShapeRange.Duplicate (издатель)"
keywords: vbapb10.chm2293780
f1_keywords: vbapb10.chm2293780
ms.prod: publisher
api_name: Publisher.ShapeRange.Duplicate
ms.assetid: e940e551-4307-aa33-5713-80f77fade8af
ms.date: 06/08/2017
ms.openlocfilehash: 9c64b5059d415944a317369432f140996648733b
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shaperangeduplicate-method-publisher"></a>Метод ShapeRange.Duplicate (издатель)

Создает копию на указанный объект **[ShapeRange](shaperange-object-publisher.md)** или **[фигуры](shape-object-publisher.md)** , добавляет новую фигуру или диапазона фигур в коллекцию **фигур** сразу же после фигуры или диапазона фигур указан изначально, а затем возвращает новый объект **ShapeRange** или **фигуры** .


## <a name="syntax"></a>Синтаксис

 _выражение_. **Дублирующиеся**

 переменная _expression_A, представляющий объект **ShapeRange** .


### <a name="return-value"></a>Возвращаемое значение

ShapeRange


## <a name="example"></a>Пример

В этом примере добавляется новый пустой странице в конце active публикации, добавляется фигура ромб на новую страницу, дублирует ромб и затем задает свойства для повторяющихся. Первый ромб будут иметь цвет заливки по умолчанию для активной цветовая схема; второй ромб будет смещение от первого и будут иметь первого контрастный цвет для активных цветовая схема.


```vb
Dim pgTemp As Page 
Dim shpTemp As Shape 
 
Set pgTemp = ActiveDocument.Pages.Add(Count:=1, After:=1) 
Set shpTemp = pgTemp.Shapes _ 
 .AddShape(Type:=msoShapeDiamond, _ 
 Left:=10, Top:=10, Width:=250, Height:=350) 
 
With shpTemp.Duplicate 
 .Left = 150 
 .Fill.ForeColor.SchemeColor = pbSchemeColorAccent1 
End With
```


