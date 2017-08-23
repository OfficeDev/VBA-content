---
title: "Метод DropCap.Clear (издатель)"
keywords: vbapb10.chm5505042
f1_keywords: vbapb10.chm5505042
ms.prod: publisher
api_name: Publisher.DropCap.Clear
ms.assetid: 7c30e774-c520-076a-41d8-7c68679f58bc
ms.date: 06/08/2017
ms.openlocfilehash: 7bcbbda0c2c2a8cc0c7461c829b4f4f39f124eed
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="dropcapclear-method-publisher"></a>Метод DropCap.Clear (издатель)

Удаляет форматирование буквицы.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Очистить**

 переменная _expression_A, представляет собой объект- **буквицу** .


## <a name="example"></a>Пример

В этом примере удаляется буквицы, форматирование в элементе frame указанный текст.


```vb
Sub ClearDropCap() 
 ActiveDocument.Pages(1).Shapes(1).TextFrame _ 
 .TextRange.DropCap.Clear 
End Sub
```


