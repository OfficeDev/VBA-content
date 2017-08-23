---
title: "Свойство TextFrame.HasText (издатель)"
keywords: vbapb10.chm3866642
f1_keywords: vbapb10.chm3866642
ms.prod: publisher
api_name: Publisher.TextFrame.HasText
ms.assetid: f8d1c660-c3f1-e835-adc3-114e6611de98
ms.date: 06/08/2017
ms.openlocfilehash: 407954fdb710f1ea431ed8b1d9b0c736250d395d
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="textframehastext-property-publisher"></a>Свойство TextFrame.HasText (издатель)

Возвращает константу **MsoTriState** , указывающее, имеет ли указанный фигуры текст, связанный с ним. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **HasText**

 переменная _expression_A, представляет собой объект- **TextFrame** .


## <a name="example"></a>Пример

Если две фигуры на первой странице активная публикация содержит текст, в этом примере Изменение размера фигуры в соответствии с текстом.


```vb
With ActiveDocument.Pages(1).Shapes(2).TextFrame 
 If .HasText Then .AutoFitText = pbTextAutoFitBestFit 
End With
```


