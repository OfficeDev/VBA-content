---
title: "Свойство Shape.Left (издатель)"
keywords: vbapb10.chm2228289
f1_keywords: vbapb10.chm2228289
ms.prod: publisher
api_name: Publisher.Shape.Left
ms.assetid: 275f5af9-9812-2a6b-bba3-704d4a7f5601
ms.date: 06/08/2017
ms.openlocfilehash: 88a62b01a647042652236e7aa94d08b7736116ce
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapeleft-property-publisher"></a>Свойство Shape.Left (издатель)

Возвращает или задает **Variant** , указывающее расстояние от левого края страницы до левого края указанного фигуры. Числовые значения находятся в точках; все остальные значения находятся в любой измерения, поддерживаемых Publisher (например, «2,5 дюйма»). Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Слева**

 переменная _expression_A, представляющий объект **фигуры** .


## <a name="example"></a>Пример

В этом примере задает горизонтальную позицию первой фигуры в активной публикации 1 дюйм от левого края страницы.


```vb
With ActiveDocument.Pages(1).Shapes(1) 
 .Left = InchesToPoints(1) 
End With
```


