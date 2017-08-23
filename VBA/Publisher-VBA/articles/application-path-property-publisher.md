---
title: "Свойство Application.Path (издатель)"
keywords: vbapb10.chm131097
f1_keywords: vbapb10.chm131097
ms.prod: publisher
api_name: Publisher.Application.Path
ms.assetid: 36ac9a9c-8235-aeba-c3d5-d39aef960cc5
ms.date: 06/08/2017
ms.openlocfilehash: 7e096ac4ba4ecb724e8867222650aa02963201e1
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="applicationpath-property-publisher"></a>Свойство Application.Path (издатель)

Возвращает **строку** , указывающее полный путь к файлу сохраненного active публикации, не включая Фамилия разделитель или файл.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Путь**

 переменная _expression_A, представляющий объект **приложения** .


## <a name="remarks"></a>Заметки

Свойство **[полное имя](document-fullname-property-publisher.md)** можно использовать для возвращения как путь и имя файла.


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


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект приложения](application-object-publisher.md)

