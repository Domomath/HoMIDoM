Imports VB = Microsoft.VisualBasic

Public Class frmMain

    Const PLCDataLengthError As Short = 99
    Const PLCDataHexError As Short = 98
    Const PLCDataWriteError As Short = 97

    Public Function FCS(ByVal Command As String) As String

        Dim i As Short
        Dim Checksum As Short

        Checksum = 0

        For i = 1 To Len(Command)
            Checksum = Asc(Mid(Command, i, 1)) Xor Checksum           
        Next i

        FCS = VB.Right("00" & Hex(Checksum), 2)

    End Function

    Public Function ReadDM(ByVal DMType As String, ByVal DMAddress As String, ByVal WordCount As Short) As String

        Dim Command As String = ""
        On Error GoTo errHandler

        If WordCount > 29 Then
            ReadDM = "Err"
            Exit Function
        End If
        Dim Type As String = ""
        Select Case DMType
            Case "Mots"
                Type = "RM"
            Case "Bits"
                Type = "RB"
        End Select


        Command = "ESC" & Type & "0" & VB.Right("0000" & DMAddress, 4) & VB.Right("00" & (Math.Truncate(WordCount)).ToString(), 2)
        Command = Command & FCS(Command) & "CR"

        If Len(Command) = 0 Then 'Timeout Occur
            ReadDM = ""
        Else
            ReadDM = Command
        End If
        Exit Function
errHandler:
        WriteLine(0, "ReadDM> " & Err.Description)
        ReadDM = ""
        Exit Function
    End Function

    Public Function WriteDM(ByVal DMType As String, ByVal DMAddress As String, ByVal WordCount As Short, ByVal Data As String) As String

        Dim Command As String = ""
        ' VBto upgrade warning: DataLen As Short	OnWrite(Integer)
        Dim DataLen As Short

        On Error GoTo errHandler

        'Check if Lenght of data is times of 4.
        DataLen = Len(Data) Mod 4
        If DataLen <> 0 Then
            WriteDM = PLCDataLengthError
            Exit Function
        End If

        If Is_Hex_Valid(Data) = 0 Then
            WriteDM = PLCDataHexError
            Exit Function
        End If

        Dim Type As String = ""
        Select Case DMType
            Case "Mots"
                Type = "RM"
            Case "Bits"
                Type = "RB"
        End Select

        Command = "ESC" & Type & "0" & VB.Right("0000" & DMAddress, 4) & VB.Right("00" & (Math.Truncate(WordCount)).ToString(), 2) & Data
        Command = Command & FCS(Command) & "CR"

        If Len(Command) = 0 Then 'Timeout Occur
            WriteDM = ""
        Else
            WriteDM = Command
        End If
        Exit Function

errHandler:
        WriteDM = PLCDataWriteError
        Exit Function

    End Function

    Public Function Is_Hex_Valid(ByVal Data As String) As Short
        Is_Hex_Valid = 0

        Dim B As Boolean
        Dim bb As Boolean '* To use for logical operations
        ' VBto upgrade warning: i As Short	OnWrite(Integer)
        Dim i As Short
        Dim j As Short '* To use in For loops
        Dim s As String = ""
        Dim values(15) As String


        s = UCase(Data)

        values(0) = "0"
        values(1) = "1"
        values(2) = "2"
        values(3) = "3"
        values(4) = "4"
        values(5) = "5"
        values(6) = "6"
        values(7) = "7"
        values(8) = "8"
        values(9) = "9"
        values(10) = "A"
        values(11) = "B"
        values(12) = "C"
        values(13) = "D"
        values(14) = "E"
        values(15) = "F"

        bb = (Len(s) > 0)

        For i = 1 To Len(s)
            B = False

            For j = 0 To 15
                B = B Or (Mid(s, i, 1) = values(j))
            Next

            bb = bb And B
        Next

        If bb Then
            Is_Hex_Valid = Len(Data)
        Else
            Is_Hex_Valid = 0
        End If

    End Function

    Public Function ParseAnswer(ByVal Answer As String) As String

        Dim Data As String = ""
        Dim Status As String = ""
        Dim Cmd As String = ""
        Dim PLCStatus As String = ""

        ' Check is the package is valid
        ' Check the header
        If VB.Left(Answer, 3) <> "ESC" Then
            ParseAnswer = "ErH" 'Header error
            Exit Function
        End If

        ' Check the tail
        If VB.Right(Answer, 2) <> "CR" Then
            ParseAnswer = "ErT" 'Tail error
            Exit Function
        End If

        Cmd = Mid(Answer, 4, 2)

        'Operation Was successful
        Select Case Cmd
            Case "RM", "RB", "SM", "SB"
                ' Answer for a read operation
                ParseAnswer = Mid(Answer, 9, Len(Answer) - 11)
                Exit Function
            Case "WM", "WB"
                'Answer for a write operation
                ParseAnswer = "00"
        End Select

        ParseAnswer = Status
    End Function

    Public Function Received(ByVal command As String) As String

        On Error GoTo errHandler

        If Len(Command) = 0 Then 'Timeout Occur
            Received = ""
        Else
            Received = ParseAnswer(command)
        End If
        Exit Function
errHandler:
        WriteLine(0, "ReadDM> " & Err.Description)
        Received = ""
        Exit Function

    End Function

End Class