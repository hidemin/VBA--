VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Matrix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private A_() As Double
Private RowsCount As Long
Private ColumnsCount As Long

Private Sub Class_Initialize()
    RowsCount = 0
    ColumnsCount = 0
End Sub

Public Sub SetDimensionOnce(RowNum As Long, ColumnNum As Long)
    If RowsCount <> 0 Then
        '一回だけだって言ってるだろ
        Stop
    End If
    
    If RowNum <= 0 Or ColumnNum <= 0 Then
        ' 行の数とか列の数は正の整数のみ許可
        Stop
    End If
    
    RowsCount = RowNum
    ColumnsCount = ColumnNum
    
    ReDim A_(RowsCount * ColumnsCount)
End Sub

Public Property Let A(RowNum As Long, ColumnNum As Long, AValue As Double)
    If RowNum < 1 Or RowsCount < RowNum Or ColumnNum < 1 Or ColumnsCount < ColumnNum Then
        Err.Raise 1024, "Matrix", "行・列の範囲を超えてる"
    End If
    
    A_(RC2I(RowNum, ColumnNum)) = AValue
End Property

Public Property Get A(RowNum As Long, ColumnNum As Long) As Double
    If RowNum < 1 Or RowsCount < RowNum Or ColumnNum < 1 Or ColumnsCount < ColumnNum Then
        Err.Raise 1024, "Matrix", "行・列の範囲を超えてる"
    End If
    
    A = A_(RC2I(RowNum, ColumnNum))
End Property

Public Property Get RowCount()
    RowCount = RowsCount
End Property

Public Property Get ColumnCount()
    ColumnCount = ColumnsCount
End Property

Public Function ToString() As String
    Dim AnsString As String
    AnsString = ""
    
    Dim RowNum As Long
    Dim ColumnNum As Long
    If RowsCount > 1 Then
        AnsString = AnsString & "{"
    End If
    For RowNum = 1 To RowsCount
        If RowNum > 1 Then
            AnsString = AnsString & ", "
        End If
        AnsString = AnsString & "{"
        For ColumnNum = 1 To ColumnsCount
            If ColumnNum > 1 Then
                AnsString = AnsString & ", "
            End If
            AnsString = AnsString & Trim(Str(A_(RC2I(RowNum, ColumnNum))))
        Next ColumnNum
        AnsString = AnsString & "}"
    Next RowNum
    If RowsCount > 1 Then
        AnsString = AnsString & "}"
    End If
    
    ToString = AnsString
End Function

Private Function RC2I(RowNum As Long, ColumnNum As Long) As Long
    RC2I = (RowNum - 1) * RowsCount + ColumnNum
End Function

Public Function Determinant() As Double
    If RowsCount <> ColumnsCount Then
        Err.Raise 1025, "Matrix", "正方行列以外は行列式だせないよう"
    End If
    
    If RowsCount = 1 Then
        Determinant = Me.A(1, 1)
        Exit Function
    End If
    
    Dim AnswerValue As Double
    AnswerValue = 0
    
    Dim ColumnNum As Long
    For ColumnNum = 1 To ColumnsCount
        Dim MinorMatrix As Matrix
        Set MinorMatrix = Me.FirstMinor(1, ColumnNum)
        AnswerValue = AnswerValue + ((-1) ^ (ColumnNum + 1)) * Me.A(1, ColumnNum) * MinorMatrix.Determinant
    Next ColumnNum
    
    Determinant = AnswerValue
End Function

Public Function Clone() As Matrix
    Dim answerMatrix As Matrix
    Set answerMatrix = New Matrix
    
    If RowsCount = 0 Then
        Set Clone = answerMatrix
        Exit Function
    End If
    
    Dim RowNum As Long
    Dim ColumnNum As Long
    For RowNum = 1 To RowsCount
        For ColumnNum = 1 To ColumnsCount
        
        Next ColumnNum
    Next RowNum
End Function

Public Function FirstMinor(MinorRowNum As Long, MinorColumnNum As Long) As Matrix
    If RowsCount = 0 Then
        Err.Raise 1025, "Matrix", "初期化前の行列から小行列を求めるのは不可では"
    End If
    If MinorRowNum < 1 Or RowsCount < MinorRowNum Or MinorColumnNum < 1 Or ColumnsCount < MinorColumnNum Then
        Err.Raise 1024, "Matrix", "行・列の範囲を超えてる"
    End If
    
    Dim answerMatrix As Matrix
    Set answerMatrix = New Matrix
    answerMatrix.SetDimensionOnce RowsCount - 1, ColumnsCount - 1
    
    Dim RowNum As Long
    Dim ColumnNum As Long
    For RowNum = 1 To RowsCount
        Dim writeRowNum As Long
        If RowNum < MinorRowNum Then
            writeRowNum = RowNum
        ElseIf RowNum = MinorRowNum Then
            writeRowNum = 0
        ElseIf RowNum > MinorRowNum Then
            writeRowNum = RowNum - 1
        End If
        For ColumnNum = 1 To ColumnsCount
            Dim writeColumnNum As Long
            If ColumnNum < MinorColumnNum Then
                writeColumnNum = ColumnNum
            ElseIf ColumnNum = MinorColumnNum Then
                writeColumnNum = 0
            ElseIf ColumnNum > MinorColumnNum Then
                writeColumnNum = ColumnNum - 1
            End If
            If writeRowNum > 0 And writeColumnNum > 0 Then
                answerMatrix.A(writeRowNum, writeColumnNum) = Me.A(RowNum, ColumnNum)
            End If
        Next ColumnNum
    Next RowNum
    
    Set FirstMinor = answerMatrix
End Function
