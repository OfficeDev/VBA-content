---
title: "Свойство ThreeDFormat.Perspective (издатель)"
keywords: vbapb10.chm3801347
f1_keywords: vbapb10.chm3801347
ms.prod: publisher
api_name: Publisher.ThreeDFormat.Perspective
ms.assetid: 5a85f7fa-2c72-e9b0-75f0-e6d6680ecd99
ms.date: 06/08/2017
ms.openlocfilehash: 44c921d2680217721cd2e73d820365617c8890d1
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="threedformatperspective-property-publisher"></a>Свойство ThreeDFormat.Perspective (издатель)

 **msoTrue** Если изменяется в Перспектива — то есть, если стенок изменяется сузить направить перспективы. **msoFalse** изменяется в случае параллельного или Ортогональная, проекции — то есть, если не сузить стенок направить перспективы. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Перспектива**

 переменная _expression_A, представляет собой объект- **ThreeDFormat** .


### <a name="return-value"></a>Возвращаемое значение

MsoTriState


## <a name="example"></a>Пример

В этом примере задается глубина объема для фигуры одно на первой странице в 100 точек и указывает, что выбирать значения параллельный или Ортогональная. В данном примере для работы указанного фигуры должен быть объемной фигуры.


```vb
Sub ChangePerspective() 
 With ActiveDocument.Pages(1).Shapes(1).ThreeD 
 .Visible = True 
 .Depth = 100 
 .Perspective = msoFalse 
 End With 
End Sub
```


