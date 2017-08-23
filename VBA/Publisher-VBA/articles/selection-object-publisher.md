---
title: "Объект Selection (издатель)"
keywords: vbapb10.chm917503
f1_keywords: vbapb10.chm917503
ms.prod: publisher
api_name: Publisher.Selection
ms.assetid: 1ebee88b-a39e-ea3a-48b0-6205621853af
ms.date: 06/08/2017
ms.openlocfilehash: c4c07102dd18279ece643845ce50bf2d46a72673
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="selection-object-publisher"></a>Объект Selection (издатель)

Представляет текущего выделения в окне или области. Выбор представляет либо выбранного (или выделенной) области в публикации или представляет указатель, если ничего не в публикации выбрано. Может быть только один объект **выбора** на панели окна публикации, и только один объект **выбора** в всего приложения может быть активным.
 


## <a name="example"></a>Пример

Свойство **[выбора](document-selection-property-publisher.md)** используется для возврата объекта **Selection** . При использовании без описатель объекта со свойством **выбора** Microsoft Publisher возвращает выборку из активной области в окне active публикации. В следующем примере копируется текущего выделения из активной публикации.
 

 

```
Sub CopySelection() 
 Selection.ShapeRange.Copy 
End Sub
```

В следующем примере определяется, какой тип элемента выбран и, если это автофигуры, заливка первую фигуру в выделение цветом. В этом примере предполагается, что имеется по крайней мере один элемент, выбранный в active pubication.
 

 



```
Sub SelectedShape() 
 If Selection.Type = pbSelectionShape Then 
 Selection.ShapeRange.Item(1).Fill.ForeColor _ 
 .RGB = RGB(Red:=200, Green:=20, Blue:=255) 
 End If 
End Sub
```

В следующем примере копирует выделение и вставляет его в первую фигуру на второй странице active публикации.
 

 



```
Sub CopyPasteSelection() 
 Selection.TextRange.Copy 
 With ActiveDocument.Pages(2).Shapes(1).TextFrame.TextRange 
 .Collapse Direction:=pbCollapseEnd 
 .InsertAfter NewText:=vbLf 
 .Paste 
 End With 
End Sub
```


## <a name="methods"></a>Методы



|**Name**|
|:-----|
|[Снятие выделения](selection-unselect-method-publisher.md)|

## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](selection-application-property-publisher.md)|
|[ChildShapeRange](selection-childshaperange-property-publisher.md)|
|[Родительский раздел](selection-parent-property-publisher.md)|
|[ShapeRange](selection-shaperange-property-publisher.md)|
|[TableCellRange](selection-tablecellrange-property-publisher.md)|
|[TextRange](selection-textrange-property-publisher.md)|
|[Type](selection-type-property-publisher.md)|

