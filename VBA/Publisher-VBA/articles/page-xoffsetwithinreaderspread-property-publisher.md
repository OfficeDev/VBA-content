---
title: "Свойство Page.XOffsetWithinReaderSpread (издатель)"
keywords: vbapb10.chm393236
f1_keywords: vbapb10.chm393236
ms.prod: publisher
api_name: Publisher.Page.XOffsetWithinReaderSpread
ms.assetid: 42ae7545-78f5-c034-33b4-f8c8f6a0b935
ms.date: 06/08/2017
ms.openlocfilehash: 2bbdf5c26e4ebb43a512d91eec8933cc94e6e4a8
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pagexoffsetwithinreaderspread-property-publisher"></a>Свойство Page.XOffsetWithinReaderSpread (издатель)

Возвращает значение типа **одного** , которое представляет расстояние (в точках) от левого края Ширина по левому краю страницы чтения. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **XOffsetWithinReaderSpread**

 переменная _expression_A, представляющий объект **Page** .


### <a name="return-value"></a>Возвращаемое значение

Один


## <a name="example"></a>Пример

В этом примере создается фигуры на страницах второго и третьего active публикации и затем задает позицию фигуры на странице третий диагонали положительно углу страницы из фигуры на второй странице. Для работы этого примера активная публикация должна иметь по крайней мере три страницы.


```vb
Sub OffsetShapePositions() 
 Dim shpOne As Shape 
 Dim intLeft As Integer 
 Dim intTop As Integer 
 Dim intWidth As Integer 
 Dim intHeight As Integer 
 
 With ActiveDocument 
 .ViewTwoPageSpread = True 
 
 With .Pages 
 intWidth = 150 
 intHeight = 150 
 intLeft = (.Item(2).Width / 2) - intWidth 
 intTop = InchesToPoints(7) 
 
 Set shpOne = .Item(2).Shapes.AddShape _ 
 (Type:=msoShape5pointStar, Left:=intLeft, _ 
 Top:=intTop, Width:=intWidth, Height:=intHeight) 
 
 intLeft = (.Item(3).XOffsetWithinReaderSpread - _ 
 .Item(2).XOffsetWithinReaderSpread) + (.Item(2) _ 
 .Width - shpOne.Left - shpOne.Width) 
 intTop = (.Item(3).YOffsetWithinReaderSpread - _ 
 .Item(2).YOffsetWithinReaderSpread) + (.Item(2) _ 
 .Height - shpOne.Top - shpOne.Height) 
 
 .Item(2).Shapes.AddShape Type:=msoShape5pointStar, _ 
 Left:=intLeft, Top:=intTop, Width:=intWidth, _ 
 Height:=intHeight 
 End With 
 End With 
End Sub
```


