---
title: "Свойство Document.FullName (издатель)"
keywords: vbapb10.chm196625
f1_keywords: vbapb10.chm196625
ms.prod: publisher
api_name: Publisher.Document.FullName
ms.assetid: 137e4310-8431-ed2a-503a-c225378a9a74
ms.date: 06/08/2017
ms.openlocfilehash: eef4c11c671940021fcc5df8010057a9868aefb1
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="documentfullname-property-publisher"></a>Свойство Document.FullName (издатель)

Возвращает **строку** , представляющую полное имя файла сохраненного active публикации, включая его путь и имя файла. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Полное имя**

 переменная _expression_A, представляющий объект **Document** .


### <a name="return-value"></a>Возвращаемое значение

String


## <a name="remarks"></a>Заметки

Свойство **полное имя** можно использовать для возвращения как путь и имя файла в виде, возвращаемом **[путь](document-path-property-publisher.md)** и **[имя](document-name-property-publisher.md)** свойства.


## <a name="example"></a>Пример

В следующем примере показано различия между **путь**, **имя**и **полное имя** свойства. В этом примере лучше всего иллюстрируется публикации при сохранении в папку по умолчанию.


```vb
Sub PathNames() 
 
 Dim strPath As String 
 Dim strName As String 
 Dim strFullName As String 
 
 strPath = Application.ActiveDocument.Path 
 strName = Application.ActiveDocument.Name 
 strFullName = Application.ActiveDocument.FullName 
 
 ' Note the file name &; path differences 
 ' while executing. 
 MsgBox "The path is: " &; strPath 
 MsgBox "The file name is: " &; strName 
 MsgBox "The path &; file name are: " &; strFullName 
 
End Sub
```


