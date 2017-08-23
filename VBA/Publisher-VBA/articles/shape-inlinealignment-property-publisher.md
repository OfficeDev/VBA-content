---
title: "Свойство Shape.InlineAlignment (издатель)"
keywords: vbapb10.chm5308694
f1_keywords: vbapb10.chm5308694
ms.prod: publisher
api_name: Publisher.Shape.InlineAlignment
ms.assetid: daef2761-2a93-25da-9c12-1fed0fdd24ab
ms.date: 06/08/2017
ms.openlocfilehash: 1a7337cfb07494e63de8dfed8462e324481495c8
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapeinlinealignment-property-publisher"></a>Свойство Shape.InlineAlignment (издатель)

Возвращает или задает **PbInlineAlignment** константа, указывающее, является ли встроенная фигура слева, справа, или выравнивание в текст. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **InlineAlignment**

 переменная _expression_A, представляющий объект **фигуры** .


## <a name="remarks"></a>Заметки

Значение свойства **InlineAlignment** может иметь одно из **[PbInlineAlignment](pbinlinealignment-enumeration-publisher.md)** константы в библиотеке типов, Microsoft Publisher.

Если фигуры еще не является встроенной, возвращается ошибка автоматизации.


## <a name="example"></a>Пример

В следующем примере второй фигура перемещается на второй странице публикации в потоке текста с помощью метода **[MoveIntoTextFlow](shape-moveintotextflow-method-publisher.md)** . Свойство **InlineAlignment** затем используется для выравнивания фигур справа.


```vb
Dim theShape As Shape 
Dim theRange As TextRange 
 
Set theRange = ActiveDocument.Pages(2).Shapes(1).TextFrame.TextRange 
Set theShape = ActiveDocument.Pages(2).Shapes(2) 
 
If Not theShape.IsInline = msoTrue Then 
 theShape.MoveIntoTextFlow Range:=theRange 
 theShape.InlineAlignment = pbInlineAlignmentRight 
End If
```


