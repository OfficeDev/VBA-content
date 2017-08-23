---
title: "Метод TabStops.ClearAll (издатель)"
keywords: vbapb10.chm5570564
f1_keywords: vbapb10.chm5570564
ms.prod: publisher
api_name: Publisher.TabStops.ClearAll
ms.assetid: bb7e2a0e-c044-872d-aa74-2683886e77a6
ms.date: 06/08/2017
ms.openlocfilehash: 4db6854c0ddd4a53cc8f2271406d809266d65397
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="tabstopsclearall-method-publisher"></a>Метод TabStops.ClearAll (издатель)

Удаляет все настраиваемые табуляции из указанного абзацев.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ClearAll**

 переменная _expression_A, представляет собой объект- **TabStops** .


## <a name="remarks"></a>Заметки

Чтобы удалить отдельные позиции табуляции, используйте метод **[снимите](tabstop-clear-method-publisher.md)** объекта **[TabStop](tabstop-object-publisher.md)** . Метод **ClearAll** не снимите флажок по умолчанию табуляции. Для работы с вкладка по умолчанию останавливается, используйте свойство **[DefaultTabStop](document-defaulttabstop-property-publisher.md)** для документа.


## <a name="example"></a>Пример

В этом примере очищается все точки пользовательской вкладки в первую фигуру на первой странице active публикации. Предполагается, что указанные форму фрагмент текста и не другого типа фигуры.


```vb
Sub ClearAllTabStops() 
 ActiveDocument.Pages(1).Shapes(1).TextFrame _ 
 .TextRange.ParagraphFormat.Tabs.ClearAll 
End Sub
```


