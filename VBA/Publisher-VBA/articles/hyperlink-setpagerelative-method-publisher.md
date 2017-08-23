---
title: "Метод Hyperlink.SetPageRelative (издатель)"
keywords: vbapb10.chm4587542
f1_keywords: vbapb10.chm4587542
ms.prod: publisher
api_name: Publisher.Hyperlink.SetPageRelative
ms.assetid: 4b2f2e84-09ce-cef6-6f22-b82642cc71fe
ms.date: 06/08/2017
ms.openlocfilehash: 5ac1c463e7a633f628f6fcfadf53bbfdf117f1df
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="hyperlinksetpagerelative-method-publisher"></a>Метод Hyperlink.SetPageRelative (издатель)

Задает тип объекта для указанного гиперссылки.


## <a name="syntax"></a>Синтаксис

 _выражение_. **SetPageRelative** ( **_RelativePage_**)

 переменная _expression_A, представляющий объект **гиперссылки** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|RelativePage|Обязательное свойство.| **PbHlinkTargetType**|Тип объекта гиперссылки.|

## <a name="remarks"></a>Заметки

Параметр RelativePage может иметь одно из следующих **PbHlinkTargetType** константы, описанные в библиотеке типов, Microsoft Publisher.



| **pbHlinkTargetTypeEmail**|| **pbHlinkTargetTypeFirstPage**|| **pbHlinkTargetTypeLastPage**|| **pbHlinkTargetTypeNextPage**|| **pbHlinkTargetTypeNone**|| **pbHlinkTargetTypePageID**|| **pbHlinkTargetTypePreviousPage**|| **pbHlinkTargetTypeURL**|

## <a name="example"></a>Пример

В следующем примере добавляет четыре новых гиперссылок фигуры одно на странице один из активных публикации и устанавливает их целевые значения соответствующим образом.


```vb
Sub SetHyperlinkRelativeTarget() 
 Dim hypNew As Hyperlink 
 Dim txtRng As TextRange 
 
 ActiveDocument.Pages(1).Shapes _ 
 .AddTextbox Orientation:=pbTextOrientationHorizontal, _ 
 Left:=10, Top:=10, Width:=200, Height:=200 
 
 Set txtRng = ActiveDocument.Pages(1).Shapes(1) _ 
 .TextFrame.TextRange 
 
 txtRng.Text = "First Page" &; vbCrLf 
 
 Set txtRng = ActiveDocument.Pages(1).Shapes(1) _ 
 .TextFrame.TextRange 
 Set hypNew = ActiveDocument.Pages(1).Shapes(1).TextFrame _ 
 .TextRange.Hyperlinks.Add(Text:=txtRng, _ 
 Address:="http://www.tailspintoys.com/") 
 
 'Change hyperlink to be a Page-relative link 
 hypNew.SetPageRelative RelativePage:=pbHlinkTargetTypeFirstPage 
 
 txtRng.Collapse pbCollapseEnd 
 txtRng.Text = "Previous Page" &; vbCrLf 
 
 Set hypNew = ActiveDocument.Pages(1).Shapes(1).TextFrame _ 
 .TextRange.Hyperlinks.Add(Text:=txtRng, _ 
 Address:="http://www.tailspintoys.com/") 
 
 hypNew.SetPageRelative RelativePage:=pbHlinkTargetTypePreviousPage 
 
 txtRng.Collapse pbCollapseEnd 
 txtRng.Text = "Next Page" &; vbCrLf 
 Set hypNew = ActiveDocument.Pages(1).Shapes(1) _ 
 .TextFrame.TextRange.Hyperlinks.Add(Text:=txtRng, _ 
 Address:="http://www.tailspintoys.com/") 
 hypNew.SetPageRelative RelativePage:=pbHlinkTargetTypeNextPage 
 
 txtRng.Collapse pbCollapseEnd 
 txtRng.Text = "Last Page" &; vbCrLf 
 Set hypNew = ActiveDocument.Pages(1).Shapes(1) _ 
 .TextFrame.TextRange.Hyperlinks.Add(Text:=txtRng, _ 
 Address:="http://www.tailspintoys.com/") 
 hypNew.SetPageRelative RelativePage:=pbHlinkTargetTypeLastPage 
 
End Sub
```


