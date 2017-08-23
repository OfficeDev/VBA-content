---
title: "Свойство PageBackground.Fill (издатель)"
keywords: vbapb10.chm8126467
f1_keywords: vbapb10.chm8126467
ms.prod: publisher
api_name: Publisher.PageBackground.Fill
ms.assetid: bb5226aa-0b47-0d0f-1310-5abb34999910
ms.date: 06/08/2017
ms.openlocfilehash: 27022cfb8caaa950f8ae381d9fb0cd2ee8d98c29
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pagebackgroundfill-property-publisher"></a>Свойство PageBackground.Fill (издатель)

 Возвращает объект **[FillFormat](fillformat-object-publisher.md)** , представляющий заливки для указанной ячейке фигуры или таблицу.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Заполните поля**

 переменная _expression_A, представляет собой объект- **PageBackground** .


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


