---
title: "Свойство значение LayoutGuides.Rows (издатель)"
keywords: vbapb10.chm1114120
f1_keywords: vbapb10.chm1114120
ms.prod: publisher
api_name: Publisher.LayoutGuides.Rows
ms.assetid: a42286ef-d955-c39d-49a4-b0e54b4d1cec
ms.date: 06/08/2017
ms.openlocfilehash: da7fe40d7cd94867bd1c1306f2b05356ae883e13
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="layoutguidesrows-property-publisher"></a>Свойство значение LayoutGuides.Rows (издатель)

Задает или возвращает значение типа **Long** , представляющее количество строк в руководстве макета. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Строк**

 переменная _expression_A, представляет собой объект- **LayoutGuides** .


## <a name="example"></a>Пример

В этом примере задается столбцов и строк руководства макет.


```vb
Sub SetLayoutGuides() 
 With ActiveDocument.LayoutGuides 
 .Columns 
 .Rows 
 End With 
End Sub
```


