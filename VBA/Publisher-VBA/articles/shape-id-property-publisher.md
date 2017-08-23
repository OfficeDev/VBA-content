---
title: "Свойство Shape.ID (издатель)"
keywords: vbapb10.chm2228325
f1_keywords: vbapb10.chm2228325
ms.prod: publisher
api_name: Publisher.Shape.ID
ms.assetid: df4ccd93-e3fa-eeef-b5ea-e99aa0dde199
ms.date: 06/08/2017
ms.openlocfilehash: f4fefd3cdfd118062d602d37fb9407e5deef90d0
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapeid-property-publisher"></a>Свойство Shape.ID (издатель)

Возвращает значение типа **Long** , представляющее тип фигуры, диапазона фигур или свойство, тип или значение мастера. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Идентификатор**

 переменная _expression_A, представляющий объект **фигуры** .


## <a name="example"></a>Пример

В этом примере тип для каждой фигуры на первой странице active публикации.


```vb
Sub ShapeID() 
 Dim shp As Shape 
 For Each shp In ActiveDocument.Pages(1).Shapes 
 MsgBox shp.ID 
 Next shp 
End Sub
```


