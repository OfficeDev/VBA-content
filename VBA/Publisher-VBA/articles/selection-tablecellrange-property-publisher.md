---
title: "Свойство Selection.TableCellRange (издатель)"
keywords: vbapb10.chm851975
f1_keywords: vbapb10.chm851975
ms.prod: publisher
api_name: Publisher.Selection.TableCellRange
ms.assetid: d683e830-6bcd-4b53-844b-605fab184a4c
ms.date: 06/08/2017
ms.openlocfilehash: a45eff4600b8acd611e1161d4b935d1b54e2d53b
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="selectiontablecellrange-property-publisher"></a>Свойство Selection.TableCellRange (издатель)

Возвращает объект **CellRange** , представляющий ячеек в таблице выбора.


## <a name="syntax"></a>Синтаксис

 _выражение_. **TableCellRange**

 переменная _expression_A, представляющий объект **Selection** .


### <a name="return-value"></a>Возвращаемое значение

CellRange


## <a name="example"></a>Пример

В этом примере заполняет ячейки таблицы в выделение.


```vb
Sub FillTableCellRange() 
 Dim intCount As Integer 
 With Selection 
 If .Type = pbSelectionTableCells Then 
 With .TableCellRange 
 For intCount = 1 To .Count 
 .Item(intCount).Fill.ForeColor.RGB = RGB _ 
 (Red:=0, Green:=255, Blue:=255) 
 Next intCount 
 End With 
 End If 
 End With 
End Sub
```


