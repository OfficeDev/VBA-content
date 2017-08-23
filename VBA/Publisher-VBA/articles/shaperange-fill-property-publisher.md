---
title: "Свойство ShapeRange.Fill (издатель)"
keywords: vbapb10.chm2293815
f1_keywords: vbapb10.chm2293815
ms.prod: publisher
api_name: Publisher.ShapeRange.Fill
ms.assetid: cdff2b6f-52f5-3ab3-c57a-4647888cd96f
ms.date: 06/08/2017
ms.openlocfilehash: 3cab4fd268bcbf5d86183e95ca579ca19a1925a1
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shaperangefill-property-publisher"></a>Свойство ShapeRange.Fill (издатель)

 Возвращает объект **[FillFormat](fillformat-object-publisher.md)** , представляющий заливки для указанной ячейке фигуры или таблицу.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Заполните поля**

 переменная _expression_A, представляющий объект **ShapeRange** .


## <a name="example"></a>Пример

В этом примере создается новый объект **автофигуры** и заполняет фигуры с зеленой.


```vb
Sub NewShapeItem() 
 
 Dim shpHeart As Shape 
 
 Set shpHeart = ThisDocument.MasterPages.Item(1).Shapes _ 
 .AddShape(Type:=msoShapeHeart, Left:=40, Top:=80, _ 
 Width:=50, Height:=50) 
 shpHeart.Fill.ForeColor.RGB = RGB(Red:=0, Green:=255, Blue:=0) 
 
End Sub
```


