---
title: "Свойство Story.HasTable (издатель)"
keywords: vbapb10.chm5832707
f1_keywords: vbapb10.chm5832707
ms.prod: publisher
api_name: Publisher.Story.HasTable
ms.assetid: bc4912e2-f521-c6b5-b5a6-a49952014966
ms.date: 06/08/2017
ms.openlocfilehash: 320ba59d77b5e43d688cdeb995cb1dae855a43d7
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="storyhastable-property-publisher"></a>Свойство Story.HasTable (издатель)

Возвращает **msoTrue** , если фигуры представляет объект **TableFrame** или **msoFalse** , если фигуры представляет любой другой тип объекта. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **HasTable**

 переменная _expression_A, представляет собой объект- **материала** .


## <a name="example"></a>Пример

В этом примере проверяется выбранной фигуре ли таблица. Если он установлен, код задает ширину столбцов один к одному дюйма (72 точки).


```vb
Sub IsTable() 
 
 With Application.Selection.ShapeRange 
 If .HasTable = msoTrue Then 
 .Table.Columns(1).Width = 72 
 End If 
 End With 
 
End Sub
```


