VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringUtil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'-----------------------
' auth : t.ichi
'-----------------------
'■グローバルデフォルトインスタンス■
'
'「Attribute VB_PredeclaredId = true」が設定されています。
'
' 移植時はエディタからコピペではなく、
' エクスポート→テキストエディタでAttributeを確認→インポートの手順を行ってください。
'

'==============================================================================
'NULLもしくは空文字である場合
'==============================================================================
Public Function IsBlank(ByVal v As Variant) As Boolean
    IsBlank = True
    If IsEmpty(v) Then: Exit Function
    If Nz(v, "") = "" Then: Exit Function
    IsBlank = False
End Function
Public Function IsNotBlank(ByVal v As Variant) As Boolean
    IsNotBlank = Not Me.IsBlank(v)
End Function

'==============================================================================
'Variantを強制的に文字型へ変換する
'==============================================================================
Public Function Equal(ByVal v1 As Variant, ByVal v2 As Variant) As Boolean
    Dim str1 As String: str1 = Me.ToString(v1)
    Dim str2 As String: str2 = Me.ToString(v2)
    Equal = (str1 = str2)
End Function

'==============================================================================
'Variantを強制的に文字型へ変換する
'==============================================================================
Public Function ToString(ByVal v As Variant, Optional ByVal nullVal As String = "") As String
    If Me.IsBlank(v) Then
        ToString = nullVal
        Exit Function
    End If
    ToString = CStr(v)
End Function

'==============================================================================
'Int型へ変換する、失敗した場合はnullValを返す
'==============================================================================
Public Function ToInt(ByVal v As Variant, Optional ByVal nullVal As Integer = 0) As Integer
On Error GoTo ex:
    Dim str As String: str = Me.ToString(v)
    If str = "" Then
        ToInt = nullVal
        Exit Function
    End If
    ToInt = CInt(str)
    Exit Function
ex:
    ToInt = nullVal
End Function

'==============================================================================
'Long型へ変換する、失敗した場合はnullValを返す
'==============================================================================
Public Function ToLong(ByVal v As Variant, Optional ByVal nullVal As Long = 0) As Long
On Error GoTo ex:
    Dim str As String: str = Me.ToString(v)
    If str = "" Then
        ToLong = nullVal
        Exit Function
    End If
    ToLong = CLng(str)
    Exit Function
ex:
    ToLong = nullVal
End Function

'==============================================================================
'Double型へ変換する、失敗した場合はnullValを返す
'==============================================================================
Public Function ToDouble(ByVal v As Variant, Optional ByVal nullVal As Double = 0) As Double
On Error GoTo ex:
    Dim str As String: str = Me.ToString(v)
    If str = "" Then
        ToDouble = nullVal
        Exit Function
    End If
    ToDouble = CDbl(str)
    Exit Function
ex:
    ToDouble = nullVal
End Function

'==============================================================================
'Date型へ変換する、失敗した場合はnullValをDate型へ変換して返す
'==============================================================================
Public Function ToDate(ByVal v As Variant, Optional ByVal valueFormat As String = "yyyy/MM/dd") As Date
    Dim str As String: str = Me.ToString(v)
    If str = "" Then
        ToDate = Date
        Exit Function
    End If
    
    str = Format(str, valueFormat)
    
    ToDate = CDate(str)
    Exit Function
End Function

'==============================================================================
'Trim
'==============================================================================
Public Function TrimLeft(ByVal v As Variant, Optional ByVal trimVal As String = " ") As String
    Dim str As String: str = Me.ToString(v)
    If Left(str, Len(trimVal)) = trimVal Then
        TrimLeft = Right(str, Len(str) - Len(trimVal))
        Exit Function
    End If
    TrimLeft = str
End Function
'==============================================================================
'Trim
'==============================================================================
Public Function TrimRight(ByVal v As Variant, Optional ByVal trimVal As String = " ") As String
    Dim str As String: str = Me.ToString(v)
    If Right(str, Len(trimVal)) = trimVal Then
        TrimRight = Left(str, Len(str) - Len(trimVal))
        Exit Function
    End If
    TrimRight = str
End Function

'===========================================================
' 日付チェック
' 西暦 YYYYMMDD形式
'===========================================================
Public Function isYMD(ByVal v As Variant) As Boolean
On Error GoTo ex
    isYMD = False
    
    Dim str As String: str = Me.ToString(v)
    If str = "" Then
        isYMD = True
        Exit Function
    End If
    
    'フォーマットチェック
    If Not Me.IsNumber(str) Then Exit Function
    If Not Me.Length(str) = 8 Then Exit Function
    
    Dim yyyy As String
    Dim mm As String
    Dim dd As String
    
    yyyy = Mid(str, 1, 4)
    mm = Mid(str, 5, 2)
    dd = Mid(str, 7, 2)

    '日付形式ではない場合
    If Not IsDate(yyyy & "/" & mm & "/" & dd) Then
        Exit Function
    End If
    
    If CInt(yyyy) < 1900 Then
        Exit Function
    End If
    
    isYMD = True
ex:
End Function

'==============================================================================
'正規表現のパターンにマッチするかどうか
'==============================================================================
Public Function IsMatch(ByVal v As Variant, ByVal pattern As String) As Boolean
    IsMatch = False
    Dim str As String: str = Me.ToString(v)
    If str = "" And pattern <> "" Then: Exit Function
    
    Dim reg, match As Object
    Set reg = CreateObject("VBScript.RegExp")
    reg.Global = True
    reg.IgnoreCase = True
    reg.pattern = pattern
    Set match = reg.Execute(v)
    If match.Count > 0 Then
        IsMatch = True
    End If
    Set match = Nothing
    Set reg = Nothing
End Function

'==============================================================================
'アルファベットのみであるか
'==============================================================================
Public Function IsAlpha(ByVal v As Variant) As Boolean
    IsAlpha = Me.IsMatch(v, "^[a-zA-Z]+$")
End Function
'==============================================================================
'大文字アルファベットのみであるか
'==============================================================================
Public Function IsAlphaUpper(ByVal v As Variant) As Boolean
    IsAlphaUpper = Me.IsMatch(v, "^[A-Z]+$")
End Function
'==============================================================================
'小文字アルファベットのみであるか
'==============================================================================
Public Function IsAlphaLower(ByVal v As Variant) As Boolean
    IsAlphaLower = Me.IsMatch(v, "^[a-z]+$")
End Function

'==============================================================================
'数字のみであるか
'==============================================================================
Public Function IsNumber(ByVal v As Variant) As Boolean
    IsNumber = Me.IsMatch(v, "^[0-9]+$")
End Function
'==============================================================================
'アルファベットor数字のみであるか
'==============================================================================
Public Function IsAlphaNumber(ByVal v As Variant) As Boolean
    IsAlphaNumber = Me.IsMatch(v, "^[a-zA-Z0-9]+$")
End Function
'==============================================================================
'大文字アルファベットor数字のみであるか
'==============================================================================
Public Function IsAlphaUpperNumber(ByVal v As Variant) As Boolean
    IsAlphaUpperNumber = Me.IsMatch(v, "^[A-Z0-9]+$")
End Function
'==============================================================================
'小文字アルファベットor数字のみであるか
'==============================================================================
Public Function IsAlphaLowerNumber(ByVal v As Variant) As Boolean
    IsAlphaLowerNumber = Me.IsMatch(v, "^[a-z0-9]+$")
End Function


'==============================================================================
'バイトカウント
'==============================================================================
Public Function ByteCount(ByVal v As Variant) As Long
    ByteCount = LenB(Me.ToString(v))
End Function

'==============================================================================
'文字数カウント
'==============================================================================
Public Function Length(ByVal v As Variant) As Long
    Length = Len(Me.ToString(v))
End Function

'==============================================================================
'左埋め
'==============================================================================
Public Function RPad(ByVal v As Variant, ByVal padLength As Integer, Optional padChr As String = " ") As String
    Dim i As Long
    Dim str As String: str = Me.ToString(v)
    If Len(str) < padLength Then
        Do While Len(str) <= padLength
            str = str & padChr
        Loop
        str = Left(str, padLength)
    End If
    RPad = str
End Function
'==============================================================================
'右埋め
'==============================================================================
Public Function LPad(ByVal v As Variant, ByVal padLength As Integer, Optional padChr As String = " ") As String
    Dim i As Long
    Dim str As String: str = Me.ToString(v)
    If Len(str) < padLength Then
        Do While Len(str) <= padLength
            str = padChr & str
        Loop
        str = Right(str, padLength)
    End If
    LPad = str
End Function

'==============================================================================
'フォーマット：数字（カンマ）
'==============================================================================
Public Function FormatNumber(ByVal v As Variant) As String
    FormatNumber = Format(Me.ToString(v), "#.##0")
End Function

'==============================================================================
'デリミタの値で文字を分割し、Array形式で返す
'==============================================================================
Public Function SpritToArray(ByVal v As Variant, ParamArray delimiters() As Variant) As String()
    Dim str As String: str = Me.ToString(v)
    Dim i As Long
    For i = LBound(delimiters) To UBound(delimiters)
        str = Replace(str, Me.ToString(delimiters(i)), Chr(0))
    Next i
    
    SpritToArray = Split(str, Chr(0))
End Function

'==============================================================================
'デリミタの値で文字を分割し、Collection形式で返す
'==============================================================================
Public Function SpritToCollection(ByVal v As Variant, ParamArray delimiters() As Variant) As Collection
    Dim i As Long
    Dim col As New Collection
    Dim ary() As String
    
    Dim str As String: str = Me.ToString(v)
    For i = LBound(delimiters) To UBound(delimiters)
        str = Replace(str, Me.ToString(delimiters(i)), Chr(0))
    Next i
    ary = Split(str, Chr(0))
    If UBound(ary) >= 0 Then
        For i = LBound(ary) To UBound(ary)
            Call col.Add(ary(i))
        Next i
    End If
    Set SpritToCollection = col
End Function

'==============================================================================
'文字を結合して返す。
'==============================================================================
Public Function Join(ParamArray args() As Variant) As String
    Dim str As String
    Dim i As Long
    str = ""
    For i = LBound(args) To UBound(args)
        str = str & Me.ToString(args(i))
    Next i
    Join = str
End Function


'==============================================================================
'配列の文字を結合して返す。デリミタが必要な場合は引数を指定
'==============================================================================
Public Function ImplodeArray(ByRef ary() As String, Optional ByVal DELIMITER As String = ",") As String
On Error GoTo ex:
    Dim i As Long
    Dim str As String: str = ""
    If UBound(ary) >= 0 Then
        For i = LBound(ary) To UBound(ary)
            If str <> "" Then: str = str & DELIMITER
            str = str & ary(i)
        Next i
    End If
ex:
    ImplodeArray = str
End Function

'==============================================================================
'Colletionを結合して返す。デリミタが必要な場合は引数を指定
'==============================================================================
Public Function ImplodeCollection(ByRef col As Collection, Optional ByVal DELIMITER As String = ",") As String
On Error GoTo ex:
    Dim v As Variant
    Dim i As Long
    Dim str As String: str = ""
    For Each v In col
        If str <> "" Then: str = str & DELIMITER
        str = str & ToString(v)
    Next
ex:
    ImplodeCollection = str
End Function

'==============================================================================
'配列の文字を結合して返す。デリミタが必要な場合は引数を指定
'==============================================================================
Public Function ToArray(ParamArray args() As Variant) As String()
    Dim ary() As String
    Dim i As Long
    ReDim ary(UBound(args))
    For i = LBound(args) To UBound(args)
        ary(i) = Me.ToString(args(i))
    Next i
    ToArray = ary
End Function
