---
title: "Свойство WizardValue.ID (издатель)"
keywords: vbapb10.chm2097155
f1_keywords: vbapb10.chm2097155
ms.prod: publisher
api_name: Publisher.WizardValue.ID
ms.assetid: d8d1ec6b-e2e7-8729-b4d2-a62a578ead11
ms.date: 06/08/2017
ms.openlocfilehash: ee81949a882d1c9668f789a7475128e6c27d38d6
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="wizardvalueid-property-publisher"></a>Свойство WizardValue.ID (издатель)

Возвращает значение типа **Long** , представляющее тип фигуры, диапазона фигур или свойство, тип или значение мастера. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Идентификатор**

 переменная _expression_A, представляет собой объект- **WizardValue** .


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


