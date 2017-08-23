---
title: "Свойство Shape.Top (издатель)"
keywords: vbapb10.chm2228306
f1_keywords: vbapb10.chm2228306
ms.prod: publisher
api_name: Publisher.Shape.Top
ms.assetid: 76ab84a9-651c-ddc6-6f7f-f98e2b71074f
ms.date: 06/08/2017
ms.openlocfilehash: 07daf8348d42f803519d555cc07d26cd46cdb6fc
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapetop-property-publisher"></a>Свойство Shape.Top (издатель)

Возвращает или задает **Variant** , который представляет расстояние между верхней части страницы и в верхней части фигуры. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **В начало**

 переменная _expression_A, представляющий объект **фигуры** .


## <a name="example"></a>Пример

В этом примере изменяется позицию, размер и тип фигуры первую фигуру на первой странице active публикации. В этом примере предполагается, что имеется по крайней мере один фигуры на первой странице active публикации.


```vb
Sub MoveSizeChangeShape() 
 With ActiveDocument.Pages(1).Shapes(1) 
 .Top = 72 
 .Left = 72 
 .Width = 150 
 .Height = 150 
 .AutoShapeType = msoShape5pointStar 
 End With 
End Sub
```


