Imports System.IO
Imports System.Text
Imports System.Reflection
Public Class Form1

    Private Structure TargetRecord
        Public AccountCode As String
        Public Name As String
        Public Type As String
        Public OpenDate As String
        Public Currency As String
    End Structure

#Region "TextAndStringFunctions"
    Public Function ReadTextFile(ByVal FName As String, Optional ByVal _UTFEncoding As Boolean = False) As List(Of String)
        Dim SL As New List(Of String)
        Dim Enc As Encoding = Encoding.Default

        Select Case _UTFEncoding
            Case False
                Enc = Encoding.Default
            Case True
                Enc = Encoding.UTF8
        End Select

        If File.Exists(FName) Then
            Dim SR As StreamReader = New StreamReader(FName, Enc)
            Do While SR.Peek() >= 0
                SL.Add(SR.ReadLine)
            Loop
            SR.Close()
        Else
            MsgBox(FName & " not found!")
        End If
        Enc = Nothing
        ReadTextFile = SL
        SL = Nothing
    End Function

    Public Function ParseStr(ByVal TS As String, ByVal WhichOne As Integer, Optional ByVal Delimiter As String = ";") As String
        Dim I As Integer
        Dim l As Integer
        Dim Cnt As Integer
        Dim _Start As Integer
        Dim _End As Integer
        Dim S As String

        Cnt = 0
        S = TS
        If Trim(S).Last <> Delimiter Then
            S = S & Delimiter
        End If

        l = Len(S)
        For I = 1 To l
            If Mid$(S, I, 1) = Delimiter Then
                Cnt = Cnt + 1
            End If
        Next I
        If Cnt < WhichOne Then
            ParseStr = ""
            Exit Function
        End If

        Cnt = 0
        I = 1
        Do While Cnt < WhichOne
            If Mid$(S, I, 1) = Delimiter Then
                Cnt = Cnt + 1
                If Cnt < WhichOne Then
                    _Start = I + 1
                End If
            End If
            I = I + 1
        Loop
        If _Start = 0 Then _Start = 1
        _End = I - 1
        ParseStr = Mid$(S, _Start, _End - _Start)
    End Function

    Function FindInStringList(ByVal SL As List(Of String), ByVal Key As String, ByVal LookUpItem As Integer, ByVal ReturnItem As Integer) As String
        Dim Result As String = ""
        Dim LItm As String

        For I = 0 To SL.Count - 1
            LItm = Trim(ParseStr(SL(I), LookUpItem))

            If LItm = Key Then
                Result = Trim(ParseStr(SL(I), ReturnItem))
                Exit For
            End If
        Next I

        Return Result
    End Function

#End Region

    Private Sub ProcessFiles(ByVal File1 As String, ByVal File2 As String)
        Dim SL1 As New List(Of String)
        Dim SL2 As New List(Of String)
        Dim Tbl As New List(Of TargetRecord)
        Dim LocalRecord As TargetRecord
        Dim StrLine As String = ""
        Dim TargetCSV As String = ""

        SL1 = ReadTextFile(File1)
        SL2 = ReadTextFile(File2)

        For I = 1 To SL1.Count - 1
            LocalRecord = New TargetRecord
            StrLine = Trim(SL1(I))
            If StrLine <> "" Then
                LocalRecord.AccountCode = ParseStr(ParseStr(StrLine, 1), 2, "|")
                LocalRecord.Name = ParseStr(StrLine, 2)
                LocalRecord.OpenDate = ParseStr(StrLine, 4)
                LocalRecord.Currency = ParseStr(StrLine, 5)
                LocalRecord.Type = FindInStringList(SL2, LocalRecord.Name, 1, 2)
                Tbl.Add(LocalRecord)
            End If
            LocalRecord = Nothing
        Next I

        SL1 = Nothing
        SL2 = Nothing

        Dim FieldNames As FieldInfo() = LocalRecord.GetType.GetFields()

5:      For Each field As FieldInfo In FieldNames
            TargetCSV = TargetCSV & field.Name & ";"
7:      Next
        TargetCSV = TargetCSV & vbCrLf

        Dim Str As String = ""
        For I = 0 To Tbl.Count - 1
            TargetCSV = TargetCSV & Tbl(I).AccountCode & ";" & Tbl(I).Name & ";" & Tbl(I).Type & ";" & Tbl(I).OpenDate & ";" & Tbl(I).Currency & vbCrLf
        Next I

        Try
            Dim objWriter As New System.IO.StreamWriter(Application.StartupPath & "\Target.csv")
            objWriter.Write(TargetCSV)
            objWriter.Close()
            objWriter = Nothing
            Label1.Text = "Target.csv file created!"
        Catch ex As Exception
            Label1.Text = ""
        End Try

        Tbl = Nothing
    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ProcessFiles(Application.StartupPath & "\File1.csv", Application.StartupPath & "\File2.csv")
    End Sub
End Class
