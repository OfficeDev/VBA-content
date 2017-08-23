---
title: "Свойство Page.YOffsetWithinReaderSpread (издатель)"
keywords: vbapb10.chm393237
f1_keywords: vbapb10.chm393237
ms.prod: publisher
api_name: Publisher.Page.YOffsetWithinReaderSpread
ms.assetid: 765adae3-af5d-ae37-5b1c-284cce8891ca
ms.date: 06/08/2017
ms.openlocfilehash: e5aa87baf85c3d132c4b6191b0112f137ba72eb8
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pageyoffsetwithinreaderspread-property-publisher"></a>Свойство Page.YOffsetWithinReaderSpread (издатель)

Возвращает значение типа **одного** , которое представляет расстояние (в точках) от верхнего края Ширина по верхнему краю страницы чтения. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **YOffsetWithinReaderSpread**

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


