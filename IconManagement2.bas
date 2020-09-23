Attribute VB_Name = "IconManagement2"
Option Explicit
'icon sizelocated in GetString(HKEY_CURRENT_USER, "Control Panel\Desktop\WindowMetrics", "Shell Icon Size", 32)
Public Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As typSHFILEINFO, ByVal cbSizeFileInfo As Long, ByVal uFlags As Long) As Long
Public Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal hIml&, ByVal i&, ByVal hDCDest&, ByVal x&, ByVal y&, ByVal Flags&) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
    Const LVM_FIRST As Long = &H1000

Public Type typSHFILEINFO
  hIcon As Long
  iIcon As Long
  dwAttributes As Long
  szDisplayName As String * 260
  szTypeName As String * 80
End Type

Public Enum IconSize
    LargeIcon = 0
    SmallIcon = 1
End Enum

Public Const SH_USEFILEATTRIBUTES As Long = &H10
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const SHGFI_DISPLAYNAME  As Long = &H200
Public Const SHGFI_EXETYPE  As Long = &H2000
Public Const SHGFI_SYSICONINDEX  As Long = &H4000
Public Const SHGFI_SHELLICONSIZE  As Long = &H4
Public Const SHGFI_TYPENAME  As Long = &H400
Public Const SHGFI_LARGEICON  As Long = &H0
Public Const SHGFI_SMALLICON  As Long = &H1
Public Const ILD_TRANSPARENT As Long = &H1
Public Const Flags = SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE Or SH_USEFILEATTRIBUTES
Public FileInfo As typSHFILEINFO

Public Sub autosizeall(lview As ListView)
    If lview.ListItems.count > 0 Then
        Dim count As Integer
        For count = 1 To lview.ColumnHeaders.count
            AutoSizeColumnHeader lview, lview.ColumnHeaders.Item(count)
        Next
    End If
End Sub

Public Sub AutoSizeColumnHeader(lview As ListView, column As ColumnHeader, Optional ByVal SizeToHeader As Boolean = True)
    SendMessage lview.hWnd, LVM_FIRST + 30, column.Index - 1, IIf(SizeToHeader, -2, -1)
End Sub

Public Function geticonhandle(FileName As String, Size As Long) 'Gets a handle to the icon
    geticonhandle = SHGetFileInfo(FileName, FILE_ATTRIBUTE_NORMAL, FileInfo, Len(FileInfo), Flags Or Size)
End Function

Public Function drawfileicon(filetype As String, Size As IconSize, destHDC As Long, x As Long, y As Long) 'Draws the icon int the destination.hdc
    drawfileicon = ImageList_Draw(geticonhandle(filetype, Size), FileInfo.iIcon, destHDC, x, y, ILD_TRANSPARENT)
End Function

Public Function HasUniqueIcon(FileName As String) As Boolean
    HasUniqueIcon = GetDefaultIcon(GetClassname(GetExtention(FileName))) = "%1"
End Function

Public Function GetFilenoext(ByVal FileName As String) As String
    Dim temp As Long
    temp = InStrRev(FileName, "\")
    If temp > 0 Then FileName = Right(FileName, Len(FileName) - temp)
    temp = InStrRev(FileName, ".")
    If temp > 0 Then FileName = Left(FileName, temp - 1)
    GetFilenoext = FileName
End Function

Public Function GetExtention(FileName As String) As String
    Dim temp As Long
    temp = InStrRev(FileName, ".")
    If temp = 0 Then
        GetExtention = GetFilename(FileName)
    Else
        GetExtention = Right(FileName, Len(FileName) - temp)
    End If
End Function

Public Function GetFilename(FileName As String) As String
    Dim temp As Long
    temp = InStrRev(FileName, "\")
    If temp = 0 Then
        GetFilename = FileName
    Else
        GetFilename = Right(FileName, Len(FileName) - temp)
    End If
End Function

Public Function GetPath(FileName As String) As String
    If InStr(FileName, "\") > 0 Then GetPath = Left(FileName, InStrRev(FileName, "\") - 1) Else GetPath = FileName
End Function

Public Function GetIconSize(Optional Size As IconSize = LargeIcon) As Long
    If Size = LargeIcon Then GetIconSize = CLng(GetString(HKEY_CURRENT_USER, "Control Panel\Desktop\WindowMetrics", "Shell Icon Size", 32)) Else GetIconSize = 16
End Function

Public Function GetClassname(ByVal Extention As String) As String
    If Left(Extention, 1) <> "." Then Extention = "." & Extention
    GetClassname = GetString(HKEY_CLASSES_ROOT, Extention)
End Function

Public Function GetDefaultIcon(Classname As String) As String
    GetDefaultIcon = GetString(HKEY_CLASSES_ROOT, Classname & "\DefaultIcon")
End Function

Public Function IsADir(FileName As String) As Boolean
    On Error Resume Next
    If Len(FileName) > 0 Then IsADir = (GetAttr(FileName) And vbDirectory) = vbDirectory
End Function

Public Function IsLike(ByVal text As String, ByVal Expression As String) As Boolean
    Dim temp As Long, tempstr() As String
    text = LCase(text)
    Expression = LCase(Expression)
    If InStr(Expression, ";") = 0 Then
        IsLike = text Like Expression
    Else
        tempstr = Split(Expression, ";")
        For temp = 0 To UBound(tempstr)
            If text Like tempstr(temp) Then
                IsLike = True
                Exit Function
            End If
        Next
    End If
End Function

Public Function GetIndex(Key As String, IML As ImageList) As Long
    On Error Resume Next
    GetIndex = IML.ListImages.Item(ValidKey(Key)).Index
End Function

Public Function GetIcon(ByVal FileName As String, IML As ImageList, picture As PictureBox) As Long
    Dim count As Long, OldFilename As String
    OldFilename = FileName
    If HasUniqueIcon(FileName) Then  'Or (isadir(filename) And FileExists(chkdir(filename, "desktop.ini"))) Then
        count = GetIndex(FileName, IML) 'is a file type with a unique icon, or is a folder with a unique icon. search by full filename
    Else
        If IsADir(FileName) = True Then
            FileName = ".Folder" 'is a normal folder
        Else
            FileName = "." & GetExtention(FileName) 'search by extention
        End If
        count = GetIndex(FileName, IML)
    End If
    
    If count = 0 Then
        GetIcon = CreateIcon(OldFilename, FileName, IML, picture)
    Else
        GetIcon = count
    End If
End Function

Public Function CreateIcon(FileName As String, ByVal Key As String, IML As ImageList, picture As PictureBox, Optional Size As IconSize = SmallIcon) As Long
     Dim count As Long
     Key = ValidKey(Key)
     picture.Cls
     picture.ScaleHeight = GetIconSize(Size)
     picture.ScaleWidth = picture.ScaleHeight
     drawfileicon FileName, Size, picture.hDC, 0, 0
     count = IML.ListImages.count
     IML.ListImages.Add , Key, picture.Image
     CreateIcon = count + 1
End Function

Public Function chkdir(Path As String, FileName As String) As String
    chkdir = Path & IIf(Right(Path, 1) = "\", Empty, "\") & FileName
End Function

Public Function ValidKey(ByVal Key As String) As String
    ValidKey = Key
    If IsNumeric(Right(Key, 1)) Then ValidKey = Key & "."
End Function
