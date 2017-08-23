---
title: "Свойство Cell.Fill (издатель)"
keywords: vbapb10.chm5111817
f1_keywords: vbapb10.chm5111817
ms.prod: publisher
api_name: Publisher.Cell.Fill
ms.assetid: 3ff3fda8-aca7-534e-6a56-99d6a3d9e9e2
ms.date: 06/08/2017
ms.openlocfilehash: 86b5ab20b325f635ba1cf0097e09ac555ec13e2c
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="cellfill-property-publisher"></a>Свойство Cell.Fill (издатель)

 Возвращает объект **[FillFormat](fillformat-object-publisher.md)** , представляющий заливки для указанной ячейке фигуры или таблицу.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Заполните поля**

 переменная _expression_A, представляет собой объект- **ячейки** .


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


