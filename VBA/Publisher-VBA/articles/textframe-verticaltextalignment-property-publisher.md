---
title: "Свойство TextFrame.VerticalTextAlignment (издатель)"
keywords: vbapb10.chm3866660
f1_keywords: vbapb10.chm3866660
ms.prod: publisher
api_name: Publisher.TextFrame.VerticalTextAlignment
ms.assetid: cd809f00-b092-c483-fe99-2aa8043fb684
ms.date: 06/08/2017
ms.openlocfilehash: 08001f3df4002a87d09547b21e0bc342157f4f49
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="textframeverticaltextalignment-property-publisher"></a>Свойство TextFrame.VerticalTextAlignment (издатель)

Возвращает или задает значение константы **PbVerticalTextAlignmentType**, представляющий вертикальное выравнивание текста в текстовом поле. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **VerticalTextAlignment**

 переменная _expression_A, представляет собой объект- **TextFrame** .


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


