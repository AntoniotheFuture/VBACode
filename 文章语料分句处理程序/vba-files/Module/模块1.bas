Attribute VB_Name = "模块1"
Const SLen = 16
Const SSplitor = ".|!|?|;|{|}" '句子分隔符
Const Comma = ","
Const STypes = "普通句,感叹句,疑问句,分号句,标题,其他"
'const
Const ResultCName = "文章序号,句子,长度,类型,语言,以数字或符号开头,大写数量"
Const NoC = 1
Const SentenceC = 2
Const LongC = 3
Const TypeC = 4
Const LangC = 5
Const LCaseCs = "[\u0061-\u007A\u00E0-\u00FE\u0161\u00FF\u0153]"
Const UCaseCs = "[\u0041-\u005A\u00C0-\u00DE\u0160\u0178\u0152]"
Const FromSheetName = "Sheet1"
Const TargetSheetName = "Sheet2"

Dim MinLength As Integer
Dim MaxLength As Integer
Dim TypeArr As String
Dim UseNumStart As Boolean
Dim MaxUpCaseC As Integer
Dim FileSplitLine As Integer
Dim FileSaveFormat As String
Dim OutputFolder As String

Dim oRegExp As Object
Dim SentExp As Object
Dim NumExp As Object
Dim DivSentExp As Object
Dim DivChangeExp As Object
Dim FloatNumExp As Object
Dim SemicolonExp As Object
Dim oMatches As Object
Dim NoNumExp As Object
Dim RCharExp As Object
Dim SpaceExp As Object
Dim NumCommaNum As Object
Dim LcaseUcase As Object
Dim SplitorToChr10 As Object
Dim NoPunctuation As Object
Dim NameHead As Object      '人名前缀处理,替换为~
Dim LCases As Object
Dim NewLineToOne As Object
Dim FirstLcase As Object   '句首小写字母

Sub GetAtts(Optional GroupName As String)
Dim WS As Worksheet
Set WS = ThisWorkbook.Sheets("定义规则")
MinLength = 15
MaxLength = 25
TypeArr = "1,2,3,4"
UseNumStart = False
MaxUpCaseC = 6
FileSplitLine = 500
FileSaveFormat = "txt"
OutputFolder = ThisWorkbook.Path & "\" & "分隔文件"

For X = 2 To WS.UsedRange.Rows.Count
    If WS.Cells(X, 1).Value = GroupName Then
        MinLength = WS.Cells(X, 2).Value
        MaxLength = WS.Cells(X, 3).Value
        TypeArr = WS.Cells(X, 4).Value
        If WS.Cells(X, 5).Value <> 0 Then UseNumStart = True
        MaxUpCaseC = WS.Cells(X, 6).Value
        FileSplitLine = WS.Cells(X, 7).Value
        FileSaveFormat = WS.Cells(X, 8).Value
        OutputFolder = WS.Cells(X, 9).Value
        Exit For
    End If
Next
End Sub

Private Sub MakeRegs()
Set oRegExp = CreateObject("vbscript.regexp")
Set SentExp = CreateObject("vbscript.regexp")
Set NumExp = CreateObject("vbscript.regexp")
Set DivSentExp = CreateObject("vbscript.regexp")
Set DivChangeExp = CreateObject("vbscript.regexp")
Set FloatNumExp = CreateObject("vbscript.regexp")
Set SemicolonExp = CreateObject("vbscript.regexp")
Set NoNumExp = CreateObject("vbscript.regexp")
Set RCharExp = CreateObject("vbscript.regexp")
Set SpaceExp = CreateObject("vbscript.regexp")
Set NumCommaNum = CreateObject("vbscript.regexp")
Set LcaseUcase = CreateObject("vbscript.regexp")
Set SplitorToChr10 = CreateObject("vbscript.regexp")
Set NoPunctuation = CreateObject("vbscript.regexp")
Set LCases = CreateObject("vbscript.regexp")
Set NameHead = CreateObject("vbscript.regexp")
Set NewLineToOne = CreateObject("vbscript.regexp")
Set FirstLcase = CreateObject("vbscript.regexp")
DivChangeExp.Global = True
DivChangeExp.ignorecase = True

'带小数点（句号）数字转换
FloatNumExp.Global = True
FloatNumExp.ignorecase = True
FloatNumExp.Pattern = "(\d+)[\u002E](\d+)"
'数字逗号改中文逗号
NumCommaNum.Global = True
NumCommaNum.ignorecase = True
NumCommaNum.Pattern = "(\d+)[\u002C](\d+)"
'小写后接大写转换
LcaseUcase.Global = True
LcaseUcase.ignorecase = False
LcaseUcase.Pattern = LCaseCs & UCaseCs

'分隔符加换行符

'句子分隔符转换
DivChangeExp.Pattern = "[\u0009]+"  '\u201C\u201D\u201D
'分号改为逗号
SemicolonExp.Global = True
SemicolonExp.ignorecase = True
SemicolonExp.Pattern = "[\u003B]+"
'分句符
DivSentExp.Global = True
DivSentExp.ignorecase = True
DivSentExp.Pattern = "([\u0021\u002E\u003F]+)"

'多余符号删掉
oRegExp.Global = True
oRegExp.ignorecase = True
oRegExp.Pattern = "[\u000A\u0022-\u0024\u0026\u0028-\u002B\u002F\u003A\u003C-\u003E\u0040\u005B-\u0060\u00A2-\u00BE\u2013-\u2018\u2020-\u20AC\u2026\uFB01\uFFFD]+"
'多余符号删掉
SentExp.Global = True
SentExp.ignorecase = True
SentExp.Pattern = "^[\u00A0\u0020]+"
'多余空格删掉
SpaceExp.Global = True
SpaceExp.ignorecase = True
SpaceExp.Pattern = "[\u0020]+"
'去掉句首的数字
NumExp.Global = True
NumExp.ignorecase = True
NumExp.Pattern = "^[0]"
NoNumExp.Global = True
NoNumExp.ignorecase = True
NoNumExp.Pattern = "\d"
RCharExp.Global = True
RCharExp.ignorecase = True
RCharExp.Pattern = "[\u0020]+$"

'判断是否字母
NoPunctuation.Global = True
NoPunctuation.ignorecase = True
NoPunctuation.Pattern = "[\u0041-\u005A\u0061-\u007A\u00C0-\u00DE\u0160\u0178\u0152\u0061-\u007A\u00E0-\u00FE\u0161\u00FF\u0153]+"

'判断人名头

NameHead.Global = True
NameHead.ignorecase = False
NameHead.Pattern = "\u0020(Dr|Jr|No|Co)\u002E"

'计算大写的数量并删掉
LCases.Global = True
LCases.ignorecase = False
LCases.Pattern = UCaseCs

'多个换行符转一个
NewLineToOne.Global = True
NewLineToOne.ignorecase = True
NewLineToOne.Pattern = "\u000A+"

FirstLcase.Global = False
FirstLcase.ignorecase = False
FirstLcase.Pattern = "^" & LCaseCs

End Sub
Sub DivMain()
NewDiv
RemovePunctuation

End Sub

Sub Sorting()
Dim WS As Worksheet
Dim TWS As Worksheet
Dim S As String
Dim FirstChar As String
Dim TX As Long
Dim m
Dim Types
Dim Length As Integer
Dim Use As Boolean
Dim Hit As Boolean

GetAtts

Set WS = ThisWorkbook.Sheets(TargetSheetName)
Set TWS = ThisWorkbook.Sheets("result")
Types = Split(TypeArr, ",")
TX = 1
TWS.UsedRange.Delete
For X = 2 To WS.UsedRange.Rows.Count
    Length = WS.Cells(X, 3).Value
    Use = False
    If MaxLength = 0 Then
        If Length >= MinLength Then
            Use = True
        End If
    Else
        If Length >= MinLength And Length <= MaxLength Then
            Use = True
        End If
    End If
    If Use Then
        Hit = False
        For Each n In Types
            If WS.Cells(X, 4).Value = CInt(n) Then
                Hit = True
            End If
        Next
        Use = Hit
    End If
    If Use Then
        If InStr(1, WS.Cells(X, 2).Value, "-") Then
            Use = False
        End If
    End If
    If Use Then
        'If X = 24639 Then
        '    MsgBox (1)
        'End If
        If WS.Cells(X, 7).Value > MaxUpCaseC Then
            Use = False
        End If
    End If
    If Use Then
        If UseNumStart Then
            TWS.Cells(TX, 1).Value = WS.Cells(X, 2).Value
            TX = TX + 1
        Else
            If WS.Cells(X, 6).Value = 0 Then
                TWS.Cells(TX, 1).Value = WS.Cells(X, 2).Value
                TX = TX + 1
            End If
        End If
    End If
    
    
Next

End Sub

'扫描每一行（文章） 》数字小数点改成。》数字逗号改成中文逗号 》小写字母后面接大写的，中间插换行符 》两个空格、、Tab转换行符 》句子分隔符 后面加换行符 》用换行符分句 》用末尾字符判断句子类型 》符号转为空格 》多个空格变一个》去掉句首空格和符号 》

Sub NewDiv()
Dim FSheet As Worksheet
Dim TSheet As Worksheet
Dim CNameArr
Dim SentArr
Dim Article As String
Dim SplitorArr
Dim TX As Long
Dim LastChar As String
Dim m '判断大写数量
Dim c As Long
Dim t1
Dim t2
Dim t3
Dim t12
Dim t23
Dim te1, te2, te3, te4, te5, te6, te7, te8

'准备对象
Set FSheet = ThisWorkbook.Sheets(FromSheetName)
Set TSheet = ThisWorkbook.Sheets(TargetSheetName)
MakeRegs
TSheet.UsedRange.AutoFilter
TSheet.UsedRange.AutoFilter
Application.Calculation = xlCalculationManual '计算改为手动
SplitorArr = Split(SSplitor, "|")
CNameArr = Split(ResultCName, ",")
TSheet.UsedRange.Delete
For Y = 0 To UBound(CNameArr)
    TSheet.Cells(1, Y + 1).Value = CNameArr(Y)
Next
TotalRow = FSheet.UsedRange.Rows.Count
Application.ScreenUpdating = False
TX = 2
UserForm1.Show (0)
UserForm1.Caption = "分割中"
For X = 1 To FSheet.UsedRange.Rows.Count

    UserForm1.ProgressBar1.Value = 100 * (X / TotalRow)
    UserForm1.Label1.Caption = X & "/" & TotalRow
    DoEvents
    Article = FSheet.Cells(X, 1).Value
    Article = FloatNumExp.Replace(Article, "$1。$2")   '数字小数点改成。
    Article = NumCommaNum.Replace(Article, "$1，$2")   '数字逗号改中文
    Article = LcaseUcase.Replace(Article, "$1" & Chr(10) & "$2")   '小写字母后面接大写的，中间插换行符
    Article = Replace(Article, "  ", Chr(10)) '两个空格、、Tab转换行符
    Article = Replace(Article, Chr(9), Chr(10)) '两个空格、、Tab转换行符
    Article = NameHead.Replace(Article, "$1$2$3！")
    Article = NewLineToOne.Replace(Article, Chr(10)) '多个换行符转一个
    For Each S In SplitorArr        '句子分隔符 后面加换行符
        Article = Replace(Article, S, S & Chr(10))
    Next
    SentArr = Split(Article, Chr(10))
    For Each S In SentArr
        If S <> "" And Len(S) > 40 Then
            m = LCases.Replace(S, "》")
            TSheet.Cells(TX, 7).Value = UBound(Split(m, "》"))
            TSheet.Cells(TX, 1).Value = X
            On Error Resume Next
            TSheet.Cells(TX, 2).Value = S
            TX = TX + 1
        End If
    Next
Next
Application.ScreenUpdating = True
UserForm1.Hide
End Sub

'去掉标点 》多个空格变一个》恢复标点 》去掉句首空格和符号 》统计词语数量
Private Sub RemovePunctuation()
Dim WS As Worksheet
Dim S As String
Dim FirstChar As String
Dim m
Set WS = ThisWorkbook.Sheets(TargetSheetName)
UserForm1.Caption = "去除标点中"

MakeRegs
TotalRow = WS.UsedRange.Rows.Count
For X = 2 To TotalRow
    
    UserForm1.Show (0)
    UserForm1.ProgressBar1.Value = 100 * (X / TotalRow)
    UserForm1.Label1.Caption = X & "/" & TotalRow
    DoEvents
    S = WS.Cells(X, 2).Value
    '去句首空格
    S = SentExp.Replace(S, "")
    '恢复数字标点
    S = Replace(S, "。", ".")
    S = Replace(S, "，", ",")
    '恢复人名标点
    S = Replace(S, "！", ".")
    '删除多余标点
    S = oRegExp.Replace(S, " ")
    
    '去句首数字
    'S = NumExp.Replace(S, "")
    S = SpaceExp.Replace(S, " ")
    WS.Cells(X, 3).Value = UBound(Split(S, " "))
    '判断句子类型
    LastChar = Right(S, 1)
    Select Case LastChar
        Case "."
            WS.Cells(X, 4).Value = 1
        Case "?"
            WS.Cells(X, 4).Value = 2
        Case "!"
            WS.Cells(X, 4).Value = 3
        Case ";"
            WS.Cells(X, 4).Value = 4
        Case Else
            WS.Cells(X, 4).Value = 5
    End Select
    '判断句首类型
    FirstChar = Left(S, 1)
    '小写
    If FirstLcase.test(FirstChar) Then
        WS.Cells(X, 6).Value = 2
    ElseIf IsNumeric(FirstChar) Then
        WS.Cells(X, 6).Value = 1
    ElseIf NoPunctuation.test(FirstChar) Then
        WS.Cells(X, 6).Value = 0
    Else
        WS.Cells(X, 6).Value = 1
    End If
    WS.Cells(X, 2).Value = S
Next
UserForm1.Hide
End Sub


Sub totxt()
GetAtts
Dim Rows
Dim OutPutMode
Dim OutputPath
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
Const TargetSheet = "result"

Dim OutSt As ADODB.Stream
Dim BinSt As ADODB.Stream
Set OutSt = CreateObject("ADODB.Stream")
Set BinSt = CreateObject("ADODB.Stream")
    
Dim fso, fe, ts
Set fso = CreateObject("Scripting.FileSystemObject")
Dim WS2 As Worksheet
Set WS2 = ThisWorkbook.Sheets(TargetSheet)
Dim c
Dim f
Dim fullname
c = 0
f = 0

GetAtts

Rows = FileSplitLine
OutPutMode = FileSaveFormat
OutputPath = OutputFolder

If OutPutMode = "Excel" Then
    fullname = ThisWorkbook.Path & "\" & Format(f, "000") & ".xlsx"
    Dim WB As Workbook
    Dim ExApp As New Excel.Application
    ExApp.Visible = False
    Set WB = ExApp.Workbooks.Add
    Application.ScreenUpdating = False
    For X = 2 To WS2.UsedRange.Rows.Count
        If c >= Rows Then
            c = 0
        End If
        If c = 0 Then
            WB.SaveAs fullname
            WB.Close
            Set WB = Nothing
            Set WB = ExApp.Workbooks.Add
            f = f + 1
            fullname = OutputPath & "\" & Format(f, "000") & ".xlsx"
        End If
        c = c + 1
        WB.Sheets(1).Cells(c, 1) = WS2.Cells(X, 1)
    Next
    WB.SaveAs fullname
    WB.Close
    Set WB = Nothing
    Application.ScreenUpdating = True
    ExApp.Quit
Else
    OutSt.Open
    OutSt.Charset = "UTF-8"
    OutSt.Type = adTypeText
    
    
    fullname = OutputPath & "\" & Format(f, "000") & ".txt"
     OutSt.WriteText Date & Time()
    For X = 2 To WS2.UsedRange.Rows.Count
        If c >= Rows Then
            c = 0
        End If
        If c = 0 Then
            OutSt.SaveToFile fullname, 2
            OutSt.Close
            OutSt.Open
            f = f + 1
            fullname = OutputPath & "\" & Format(f, "000") & ".txt"
        End If
        OutSt.WriteText WS2.Cells(X, 1) & vbCrLf
        OutSt.SkipLine
        c = c + 1
    Next
    OutSt.SaveToFile fullname, 2
    OutSt.Close
End If
End Sub

Function ToUnicode(str As String) As String
Dim i As Long
Dim chrTmp As String
Dim ByteLower As String, ByteUpper As String
For i = 1 To Len(str)
    result = result & "\u"
    chrTmp = Mid(str, i, 1)
    ByteLower = Hex$(AscB(MidB(chrTmp, 1, 1)))
    If Len(ByteLower) = 1 Then ByteLower = "0" & ByteLower
        ByteUpper = Hex$(AscB(MidB$(chrTmp, 2, 1)))
    If Len(ByteUpper) = 1 Then ByteUpper = "0" & ByteUpper
    result = result & ByteUpper & ByteLower & ""
Next
ToUnicode = result
End Function

Private Sub ShowTables()
UserForm2.Show (0)
End Sub

Sub importfromword()
Dim Path As String
Dim WApp As New Word.Application
Dim Doc As Word.Document
Dim p As Word.Paragraph
Dim WS As Worksheet
Dim X As Long
Set FilePicker = Application.FileDialog(msoFileDialogFilePicker)
With FilePicker
    .Filters.Clear
    .Filters.Add "常规word文档", "*.docx,*.docm,*.Doc"
    .Title = "选择word文档"
    .AllowMultiSelect = False
    If .Show = -1 Then
        Path = .SelectedItems(1)
    End If
End With
Application.StatusBar = "导入中"
Set WS = ThisWorkbook.Sheets("Sheet1")
WS.Rows.Delete
WS.Cells(1, 1) = "Title"
WS.Cells(1, 2) = "Sent"
Set Doc = WApp.Documents.Open(Path)
X = 1
For Each p In Doc.Paragraphs
    X = X + 1
    WS.Cells(X, 2).Value = p.Range.Text
Next
WApp.Quit False
MsgBox ("done")

End Sub
