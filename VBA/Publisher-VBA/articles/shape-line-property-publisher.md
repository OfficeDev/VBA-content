---
title: "Свойство Shape.Line (издатель)"
keywords: vbapb10.chm2228290
f1_keywords: vbapb10.chm2228290
ms.prod: publisher
api_name: Publisher.Shape.Line
ms.assetid: 3d53f917-87ad-159d-65c3-e6fdfa72b15e
ms.date: 06/08/2017
ms.openlocfilehash: 5c7c3e1a679f4f6743546127e5aeb2e59f29bfcc
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapeline-property-publisher"></a>Свойство Shape.Line (издатель)

Возвращает объект **[LineFormat](lineformat-object-publisher.md)** , который содержит строку свойства для указанного фигуры форматирования. (Для строки, сама линия представляет объект **LineFormat** ; для фигуры с границей, границы которого представляет объект **LineFormat** .).


## <a name="syntax"></a>Синтаксис

 _выражение_. **Строка**

 переменная _expression_A, представляющий объект **фигуры** .


## <a name="example"></a>Пример

В этом примере добавляет синий пунктирная линия active публикации.


```vb
With ActiveDocument.Pages(1).Shapes _ 
 .AddLine(BeginX:=10, BeginY:=10, _ 
 EndX:=250, EndY:=250).Line 
 .DashStyle = msoLineDashDotDot 
 .ForeColor.RGB = RGB(50, 0, 128) 
End With
```

В этом примере добавляется нескольких для первой страницы и затем устанавливаются ее границу, чтобы быть 8 пунктов толстые и красным.




```vb
With ActiveDocument.Pages(1).Shapes _ 
 .AddShape(Type:=msoShapeCross, _ 
 Left:=10, Top:=10, Width:=50, Height:=70).Line 
 .Weight = 8 
 .ForeColor.RGB = RGB(255, 0, 0) 
End With
```


