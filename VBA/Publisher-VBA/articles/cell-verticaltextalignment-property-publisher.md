---
title: "Свойство Cell.VerticalTextAlignment (издатель)"
keywords: vbapb10.chm5111840
f1_keywords: vbapb10.chm5111840
ms.prod: publisher
api_name: Publisher.Cell.VerticalTextAlignment
ms.assetid: 793bf932-15d0-cce9-1d5d-aee5d260e1a0
ms.date: 06/08/2017
ms.openlocfilehash: 4488a5d660a471189034c6abb2b035bbc3ffd2df
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="cellverticaltextalignment-property-publisher"></a>Свойство Cell.VerticalTextAlignment (издатель)

Возвращает или задает значение константы **PbVerticalTextAlignmentType**, представляющий вертикальное выравнивание текста в текстовом поле. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **VerticalTextAlignment**

 переменная _expression_A, представляет собой объект- **ячейки** .


## <a name="remarks"></a>Заметки

Значение свойства **VerticalTextAlignment** может иметь одно из следующих констант **PbVerticalTextAlignmentType** .



| **pbVerticalTextAlignmentBottom**|| **pbVerticalTextAlignmentCenter**|| **pbVerticalTextAlignmentTop**|

## <a name="example"></a>Пример

В этом примере по вертикали Центрирует текст в элементе frame указанный текст. В этом примере предполагается, что имеется по крайней мере один фигуры на первой странице active публикации.


```vb
Sub SetVerticalAlignment() 
 ActiveDocument.Pages(1).Shapes(1).TextFrame _ 
 .VerticalTextAlignment = pbVerticalTextAlignmentCenter 
End Sub
```


