---
title: "Свойство Shape.Fill (издатель)"
keywords: vbapb10.chm2228279
f1_keywords: vbapb10.chm2228279
ms.prod: publisher
api_name: Publisher.Shape.Fill
ms.assetid: ff1b8d02-150e-e023-2f0a-b1608cc99644
ms.date: 06/08/2017
ms.openlocfilehash: a7ceee543df06f2689dde5e953d619009308e0b7
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapefill-property-publisher"></a>Свойство Shape.Fill (издатель)

 Возвращает объект **[FillFormat](fillformat-object-publisher.md)** , представляющий заливки для указанной ячейке фигуры или таблицу.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Заполните поля**

 переменная _expression_A, представляющий объект **фигуры** .


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


