---
title: "Свойство Shape.WebCheckBox (издатель)"
keywords: vbapb10.chm2228344
f1_keywords: vbapb10.chm2228344
ms.prod: publisher
api_name: Publisher.Shape.WebCheckBox
ms.assetid: 13796525-584f-7109-5dea-1f2baf1efda7
ms.date: 06/08/2017
ms.openlocfilehash: 52ba645910c1322f25733060b4e1062abeb157af
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapewebcheckbox-property-publisher"></a>Свойство Shape.WebCheckBox (издатель)

Возвращает объект **[WebCheckBox](webcheckbox-object-publisher.md)** , связанный с указанным фигуры.


## <a name="syntax"></a>Синтаксис

 _выражение_. **WebCheckBox**

 переменная _expression_A, представляющий объект **фигуры** .


### <a name="return-value"></a>Возвращаемое значение

WebCheckBox


## <a name="example"></a>Пример

В этом примере создается новый Web флажок и указывает, что установлен флажок состояние по умолчанию.


```vb
Dim shpNew As Shape 
Dim wcbTemp As WebCheckBox 
 
Set shpNew = ActiveDocument.Pages(1).Shapes _ 
 .AddWebControl(Type:=pbWebControlCheckBox, Left:=100, _ 
 Top:=123, Width:=17, Height:=12) 
 
Set wcbTemp = shpNew.WebCheckBox 
 
wcbTemp.Selected = msoTrue
```


