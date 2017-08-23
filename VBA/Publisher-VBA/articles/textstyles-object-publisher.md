---
title: "Объект TextStyles (издатель)"
keywords: vbapb10.chm5963775
f1_keywords: vbapb10.chm5963775
ms.prod: publisher
api_name: Publisher.TextStyles
ms.assetid: 8a250160-0400-62e7-8301-5a5743fb2485
ms.date: 06/08/2017
ms.openlocfilehash: 78465b532976e9d493cf35f982cc3db436c4c83a
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="textstyles-object-publisher"></a>Объект TextStyles (издатель)

Коллекция объектов **[стиля текста](textstyle-object-publisher.md)** , представляющих встроенные и пользовательские стили в документе.
 


## <a name="example"></a>Пример

Свойство **TextStyles** используется для возврата коллекции **TextStyles** . В следующем примере создается таблица и перечислены все стили в активной публикации.
 

 

```
Sub ListTextStyles() 
 Dim sty As TextStyle 
 Dim tbl As Table 
 Dim intRow As Integer 
 
 With ActiveDocument 
 Set tbl = .Pages(1).Shapes.AddTable(NumRows:=.TextStyles.Count, _ 
 NumColumns:=2, Left:=72, Top:=72, Width:=488, Height:=12).Table 
 For Each sty In .TextStyles 
 intRow = intRow + 1 
 With tbl.Rows(intRow) 
 .Cells(1).text = sty.Name 
 .Cells(2).text = sty.BaseStyle 
 End With 
 Next sty 
 End With 
End Sub
```

Используйте метод **[Add](textstyles-add-method-publisher.md)** создает новый стиль пользовательских и добавляет его в коллекцию **TextStyles** . В следующем примере создается новый стиль и применяется к абзац с позиции курсора.
 

 



```
Sub ApplyTextStyle() 
 Dim styNew As TextStyle 
 Dim fntStyle As Font 
 
 'Create a new style 
 Set styNew = ActiveDocument.TextStyles.Add(StyleName:="NewStyle") 
 Set fntStyle = styNew.Font 
 
 'Format the Font object 
 With fntStyle 
 .Name = "Tahoma" 
 .Size = 20 
 .Bold = msoTrue 
 End With 
 
 'Apply the Font object formatting to the new style 
 styNew.Font = fntStyle 
 
 'Apply the new style to the selected paragraph 
 Selection.TextRange.ParagraphFormat.TextStyle = "NewStyle" 
End Sub
```


## <a name="methods"></a>Методы



|**Name**|
|:-----|
|[Добавление](textstyles-add-method-publisher.md)|
|[Элемент](textstyles-item-method-publisher.md)|

## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](textstyles-application-property-publisher.md)|
|[Count](textstyles-count-property-publisher.md)|
|[Родительский раздел](textstyles-parent-property-publisher.md)|

