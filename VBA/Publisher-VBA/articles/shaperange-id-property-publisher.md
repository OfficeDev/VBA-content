---
title: "Свойство ShapeRange.ID (издатель)"
keywords: vbapb10.chm2293861
f1_keywords: vbapb10.chm2293861
ms.prod: publisher
api_name: Publisher.ShapeRange.ID
ms.assetid: d7ad646b-be40-2ac4-9d3e-faa37f8bf456
ms.date: 06/08/2017
ms.openlocfilehash: 8b8781dc516349c1386068a08570111b4786770c
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shaperangeid-property-publisher"></a>Свойство ShapeRange.ID (издатель)

Возвращает значение типа **Long** , представляющее тип фигуры, диапазона фигур или свойство, тип или значение мастера. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Идентификатор**

 переменная _expression_A, представляющий объект **ShapeRange** .


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


