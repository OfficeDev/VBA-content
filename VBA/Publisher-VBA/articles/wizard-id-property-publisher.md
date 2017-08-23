---
title: "Свойство Wizard.ID (издатель)"
keywords: vbapb10.chm1441795
f1_keywords: vbapb10.chm1441795
ms.prod: publisher
api_name: Publisher.Wizard.ID
ms.assetid: ce7df9d3-052a-6cb6-e24d-4cb5192551d0
ms.date: 06/08/2017
ms.openlocfilehash: 18ed265a0eb345cfcdc530de4844efee8ec98cbd
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="wizardid-property-publisher"></a>Свойство Wizard.ID (издатель)

Возвращает значение типа **Long** , представляющее тип фигуры, диапазона фигур или свойство, тип или значение мастера. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Идентификатор**

 переменная _expression_A, представляющий объект **мастера** .


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


