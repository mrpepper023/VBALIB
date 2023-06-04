Attribute VB_Name = "UtilZipFormat"
'=========================================================================================
'UtilZipFormat 20230527
'
'UtilZipFormatはZipアーカイブのみを受け取るアプリケーション向けの無圧縮ZIPを生成する
'=========================================================================================
'指定したファイル名配列から無圧縮ZIPアーカイブを生成する
'Public Sub MakeZip(Zip$, Files$())
'=========================================================================================
' public domain
' https://gist.github.com/7shi/573576
' 無圧縮ZIPの出力


Option Explicit

Private Crc32Table&(255)

Private Type ZipHeader
    ver As Integer
    flags As Integer
    compression As Integer
    dos_time As Integer
    dos_date As Integer
    crc32 As Long
    compressed_size As Long
    uncompressed_size As Long
    filename_length As Integer
    extra_field_length As Integer
    fname_ As String
    attrs_ As Long
    pos_ As Long
End Type

Private Sub InitCrc32Table()
    Dim i%, J%, R&, R1&
    For i = 0 To 255
        R = i
        For J = 0 To 7
            R1 = R And 1
            R = (R - R1) / 2
            If R < 0 Then R = R - &H80000000
            If R1 Then R = R Xor &HEDB88320
        Next J
        Crc32Table(i) = R
    Next i
End Sub

Public Function GetCrc32&(A$)
    Dim R&, i%, B As Byte
    If Crc32Table(255) = 0 Then InitCrc32Table
    R = Not 0
    For i = 1 To Len(A)
        B = Asc(Mid(A, i, 1))
        R = (Int(R / 256) And &HFFFFFF) Xor Crc32Table((R Xor B) And &HFF)
    Next i
    GetCrc32 = Not R
End Function

Public Function GetCrc32FromFile&(Path$)
    Dim R&, i&, B As Byte, FL&
    If Crc32Table(255) = 0 Then InitCrc32Table
    FL = FileLen(Path)
    Open Path For Binary Lock Read As #2
    R = Not 0
    For i = 1 To FL
        Get #2, , B
        R = (Int(R / 256) And &HFFFFFF) Xor Crc32Table((R Xor B) And &HFF)
    Next i
    Close #2
    GetCrc32FromFile = Not R
End Function

Private Function GetDosDate%(DT As Date)
    Dim T&
    T = ((Year(DT) - 1980) * 512 + Month(DT) * 32 + Day(DT)) And 65535
    If T >= 32768 Then T = T - 65536
    GetDosDate = T
End Function

Private Function GetDosTime%(DT As Date)
    Dim T&
    T = Hour(DT) * 2048 + Minute(DT) * 32 + Int(Second(DT) / 2)
    If T >= 32768 Then T = T - 65536
    GetDosTime = T
End Function

Private Function Path_GetFileName$(A$)
    Dim P%
    P = InStrRev(A, "\")
    If P > 0 Then
        Path_GetFileName = Mid(A, P + 1)
    Else
        Path_GetFileName = A
    End If
End Function

Private Sub WriteZipHeader(F%, ZH As ZipHeader)
    Put #F, , ZH.ver
    Put #F, , ZH.flags
    Put #F, , ZH.compression
    Put #F, , ZH.dos_time
    Put #F, , ZH.dos_date
    Put #F, , ZH.crc32
    Put #F, , ZH.compressed_size
    Put #F, , ZH.uncompressed_size
    Put #F, , ZH.filename_length
    Put #F, , ZH.extra_field_length
End Sub

Public Sub MakeZip(Zip$, Files$())
    Dim ZHS() As ZipHeader, ZHLen%, i%, J&, FL&, Path$, Name$, B As Byte
    Dim DT As Date, DS&, DL&
    
    On Error Resume Next
    Kill Zip
    On Error GoTo 0
    Open Zip For Binary Lock Write As #1
    ZHLen = UBound(Files) + 1
    ReDim ZHS(ZHLen - 1)
    
    For i = 0 To ZHLen - 1
        Path = Files(i)
        Name = Path_GetFileName(Path)
        FL = FileLen(Files(i))
        DT = FileDateTime(Files(i))
        With ZHS(i)
            .ver = 10
            .flags = 0
            .compression = 0
            .dos_time = GetDosTime(DT)
            .dos_date = GetDosDate(DT)
            .crc32 = GetCrc32FromFile(Path)
            .compressed_size = FL
            .uncompressed_size = FL
            .filename_length = LenB(StrConv(Name, vbFromUnicode))
            .extra_field_length = 0
            .fname_ = Name
            .attrs_ = GetAttr(Path)
            .pos_ = Seek(1) - 1
        End With
        Put #1, , CStr("PK" & Chr(3) & Chr(4))
        WriteZipHeader 1, ZHS(i)
        Put #1, , Name
        Open Path For Binary Lock Read As #2
        For J = 1 To FL
            Get #2, , B
            Put #1, , B
        Next J
        Close #2
    Next i
    
    DS = Seek(1) - 1
    For i = 0 To ZHLen - 1
        Put #1, , CStr("PK" & Chr(1) & Chr(2))
        Put #1, , ZHS(i).ver
        WriteZipHeader 1, ZHS(i)
        Put #1, , CInt(0)
        Put #1, , CInt(0)
        Put #1, , CInt(0)
        Put #1, , ZHS(i).attrs_
        Put #1, , ZHS(i).pos_
        Put #1, , ZHS(i).fname_
    Next i
    DL = (Seek(1) - 1) - DS
    
    Put #1, , CStr("PK" & Chr(5) & Chr(6))
    Put #1, , CInt(0)
    Put #1, , CInt(0)
    Put #1, , ZHLen
    Put #1, , ZHLen
    Put #1, , DL
    Put #1, , DS
    Put #1, , CInt(0)
    
    Close #1
End Sub

