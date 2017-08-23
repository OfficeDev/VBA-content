---
title: "Метод Shapes.SelectAll (издатель)"
keywords: vbapb10.chm2162726
f1_keywords: vbapb10.chm2162726
ms.prod: publisher
api_name: Publisher.Shapes.SelectAll
ms.assetid: 67b88529-814d-c029-1bde-e5dade87636a
ms.date: 06/08/2017
ms.openlocfilehash: ff0620bd95a9bf8b296873f5a778d4ff52c67c1b
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapesselectall-method-publisher"></a>Метод Shapes.SelectAll (издатель)

Выбирает все фигуры в определенной коллекции **[фигур](shapes-object-publisher.md)** .


## <a name="syntax"></a>Синтаксис

 _выражение_. **SelectAll**

 переменная _expression_A, представляет собой объект- **фигур** .


## <a name="example"></a>Пример

В этом примере выбирает всех фигур на странице один из активных публикации.


```vb
ActiveDocument.Pages(1).Shapes.SelectAll
```


