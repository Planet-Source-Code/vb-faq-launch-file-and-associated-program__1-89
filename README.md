<div align="center">

## Launch file and associated program


</div>

### Description

How do I launch a file in its associated program?

The Shell statement unfortunately only supports launching an EXE file directly. If you want to be able to launch, i.e. Microsoft Word by calling a .DOC file only, you can make your VB application launch the associated program with the document using the following method:
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[VB FAQ](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/vb-faq.md)
**Level**          |Unknown
**User Rating**    |4.4 (40 globes from 9 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/vb-faq-launch-file-and-associated-program__1-89/archive/master.zip)

### API Declarations

```

#IF WIN32 THEN
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
(ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd _
As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
#ELSE
Declare Function ShellExecute Lib "SHELL" (ByVal hwnd%, _
ByVal lpszOp$, ByVal lpszFile$, ByVal lpszParams$, _
ByVal lpszDir$, ByVal fsShowCmd%) As Integer
Declare Function GetDesktopWindow Lib "USER" () As Integer
#END IF
Private Const SW_SHOWNORMAL = 1
```


### Source Code

```
Function StartDoc(DocName As String) As Long
  Dim Scr_hDC As Long
  Scr_hDC = GetDesktopWindow()
  StartDoc = ShellExecute(Scr_hDC, "Open", DocName, "", "C:\", SW_SHOWNORMAL)
End Function
Private Sub Form_Click()
  Dim r As Long
  r = StartDoc("c:\my documents\word\myletter.doc")
  Debug.Print "Return code from Startdoc: "; r
End Sub
```

