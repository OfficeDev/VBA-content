---
title: "Свойство Document.ActiveView (издатель)"
keywords: vbapb10.chm196707
f1_keywords: vbapb10.chm196707
ms.prod: publisher
api_name: Publisher.Document.ActiveView
ms.assetid: 1448c8c6-30e5-2e2a-f124-ebf544d8f297
ms.date: 06/08/2017
ms.openlocfilehash: 700e725086ee9b12effd85b1d8b595bcef50b7c4
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="documentactiveview-property-publisher"></a>Свойство Document.ActiveView (издатель)

Возвращает объект **[представления](view-object-publisher.md)** , представляющее атрибуты представления для указанного документа. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ActiveView**

 переменная _expression_A, представляющий объект **Document** .


### <a name="return-value"></a>Возвращаемое значение

Представление


## <a name="example"></a>Пример

В следующем примере задается zoom active публикации на экран.


```vb
Sub SetActiveZoom() 
 Dim viewTemp As View 
 
 ActiveDocument.Pages(1).Shapes.AddShape 1, 10, 10, 50, 50 
 Set viewTemp = ActiveDocument.ActiveView 
 ActiveDocument.Pages(1).Shapes(1).Select 
 viewTemp.Zoom = pbZoomFitSelection 
End Sub
```


