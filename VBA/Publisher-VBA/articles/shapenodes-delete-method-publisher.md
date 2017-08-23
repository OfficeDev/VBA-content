---
title: "Метод ShapeNodes.Delete (издатель)"
keywords: vbapb10.chm3473425
f1_keywords: vbapb10.chm3473425
ms.prod: publisher
api_name: Publisher.ShapeNodes.Delete
ms.assetid: 09f7a8ef-cefd-5a68-f0a6-e99c2f111ea6
ms.date: 06/08/2017
ms.openlocfilehash: 047dff49a4d6de0914d0994e919bea28feeeda2b
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapenodesdelete-method-publisher"></a>Метод ShapeNodes.Delete (издатель)

Удаляет объект узел указанного фигуры.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Удаление** ( **_Индекс_**)

 переменная _expression_A, представляет собой объект- **ShapeNodes** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Индекс|Обязательное свойство.| **[INT]**| **Длинные**. Число фигур узел для удаления.|

## <a name="example"></a>Пример

В этом примере удаляется первый узел в первую фигуру в активной публикации.


```vb
Sub DeleteNode() 
 ActiveDocument.Pages(1).Shapes(1).Nodes.Delete Index:=1 
End Sub
```


