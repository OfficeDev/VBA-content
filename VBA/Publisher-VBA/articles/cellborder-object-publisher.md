---
title: "Объект CellBorder (издатель)"
keywords: vbapb10.chm5308415
f1_keywords: vbapb10.chm5308415
ms.prod: publisher
api_name: Publisher.CellBorder
ms.assetid: c4eddeac-54cd-95ff-9423-b06e515a720e
ms.date: 06/08/2017
ms.openlocfilehash: 04bd142f4f3788e06ff292eabcfd2d4ecfdfbf7f
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="cellborder-object-publisher"></a>Объект CellBorder (издатель)

Представляет параметры цвета и вес границы ячеек.
 


## <a name="example"></a>Пример

Использование различных свойств границы объекта **ячейки** для возвращения различных границ ячейки (слева, справа, верхней, нижней и вправо). В следующем примере извлекается верхняя граница первую ячейку в таблице.
 

 

```
Dim cbTemp As CellBorder 
 
Set cbTemp = ActiveDocument.Pages(1) _ 
 .Shapes(1).Table.Cells.Item(1).BorderTop
```

Использование свойства **[цвета](cellborder-color-property-publisher.md)** и **[Вес](cellborder-weight-property-publisher.md)** объекта **CellBorder** для быстрого форматирования внешний вид границы. В следующем примере создается левая граница первую ячейку в таблице красной и два аспекта толстые.
 

 



```
Dim cbTemp As CellBorder 
 
Set cbTemp = ActiveDocument.Pages(1) _ 
 .Shapes(1).Table.Cells.Item(1).BorderLeft 
 
cbTemp.Color.RGB = RGB(255, 0, 0) 
cbTemp.Weight = 2
```


## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](cellborder-application-property-publisher.md)|
|[Цвет](cellborder-color-property-publisher.md)|
|[Родительский раздел](cellborder-parent-property-publisher.md)|
|[Вес](cellborder-weight-property-publisher.md)|

