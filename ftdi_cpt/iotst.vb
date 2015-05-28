Imports System
'Imports D2XXUnit
Namespace iotst.Units
	Public Class iotst
		Public Shared Saved_Handle As Integer = 0
		Public Shared PortAIsOpen As Boolean = False
		Public Const USBBuffSize As Integer = &H4000
		' Exported Functions
		Public Shared Function HexWrdToStr(ByVal Dval As Integer) As String
			Dim result As String
			Dim i As Integer
			Dim retstr As String
			retstr = ""
            i = CInt((Dval And &HF000) / &H1000)
			Select Case i
				Case 0
					retstr = retstr & "0"c
				Case 1
					retstr = retstr & "1"c
				Case 2
					retstr = retstr & "2"c
				Case 3
					retstr = retstr & "3"c
				Case 4
					retstr = retstr & "4"c
				Case 5
					retstr = retstr & "5"c
				Case 6
					retstr = retstr & "6"c
				Case 7
					retstr = retstr & "7"c
				Case 8
					retstr = retstr & "8"c
				Case 9
					retstr = retstr & "9"c
				Case 10
					retstr = retstr & "A"c
				Case 11
					retstr = retstr & "B"c
				Case 12
					retstr = retstr & "C"c
				Case 13
					retstr = retstr & "D"c
				Case 14
					retstr = retstr & "E"c
				Case 15
					retstr = retstr & "F"c
			End Select
            i = CInt((Dval And &HF00) / &H100)
			Select Case i
				Case 0
					retstr = retstr & "0"c
				Case 1
					retstr = retstr & "1"c
				Case 2
					retstr = retstr & "2"c
				Case 3
					retstr = retstr & "3"c
				Case 4
					retstr = retstr & "4"c
				Case 5
					retstr = retstr & "5"c
				Case 6
					retstr = retstr & "6"c
				Case 7
					retstr = retstr & "7"c
				Case 8
					retstr = retstr & "8"c
				Case 9
					retstr = retstr & "9"c
				Case 10
					retstr = retstr & "A"c
				Case 11
					retstr = retstr & "B"c
				Case 12
					retstr = retstr & "C"c
				Case 13
					retstr = retstr & "D"c
				Case 14
					retstr = retstr & "E"c
				Case 15
					retstr = retstr & "F"c
			End Select
            i = CInt((Dval And &HF0) / &H10)
			Select Case i
				Case 0
					retstr = retstr & "0"c
				Case 1
					retstr = retstr & "1"c
				Case 2
					retstr = retstr & "2"c
				Case 3
					retstr = retstr & "3"c
				Case 4
					retstr = retstr & "4"c
				Case 5
					retstr = retstr & "5"c
				Case 6
					retstr = retstr & "6"c
				Case 7
					retstr = retstr & "7"c
				Case 8
					retstr = retstr & "8"c
				Case 9
					retstr = retstr & "9"c
				Case 10
					retstr = retstr & "A"c
				Case 11
					retstr = retstr & "B"c
				Case 12
					retstr = retstr & "C"c
				Case 13
					retstr = retstr & "D"c
				Case 14
					retstr = retstr & "E"c
				Case 15
					retstr = retstr & "F"c
			End Select
			i = Dval And &HF
			Select Case i
				Case 0
					retstr = retstr & "0"c
				Case 1
					retstr = retstr & "1"c
				Case 2
					retstr = retstr & "2"c
				Case 3
					retstr = retstr & "3"c
				Case 4
					retstr = retstr & "4"c
				Case 5
					retstr = retstr & "5"c
				Case 6
					retstr = retstr & "6"c
				Case 7
					retstr = retstr & "7"c
				Case 8
					retstr = retstr & "8"c
				Case 9
					retstr = retstr & "9"c
				Case 10
					retstr = retstr & "A"c
				Case 11
					retstr = retstr & "B"c
				Case 12
					retstr = retstr & "C"c
				Case 13
					retstr = retstr & "D"c
				Case 14
					retstr = retstr & "E"c
				Case 15
					retstr = retstr & "F"c
			End Select
			result = retstr
			Return result
		End Function

		Public Shared Function HexDWrdToStr(ByVal Dval As UInteger) As String
			Dim result As String
			Dim i As UInteger
			Dim tmpstr As String
            i = CUInt(Dval \ &H10000)
            tmpstr = HexWrdToStr(CInt(i))
            i = CUInt(Dval And CUInt(&HFFFF))
            tmpstr = tmpstr & HexWrdToStr(CInt(i))
			result = tmpstr
			Return result
		End Function

		Public Shared Function HexByteToStr(ByVal Dval As Integer) As String
			Dim result As String
			Dim i As Integer
			Dim retstr As String
			retstr = ""
            i = CInt((Dval And &HF0) / &H10)
			Select Case i
				Case 0
					retstr = retstr & "0"c
				Case 1
					retstr = retstr & "1"c
				Case 2
					retstr = retstr & "2"c
				Case 3
					retstr = retstr & "3"c
				Case 4
					retstr = retstr & "4"c
				Case 5
					retstr = retstr & "5"c
				Case 6
					retstr = retstr & "6"c
				Case 7
					retstr = retstr & "7"c
				Case 8
					retstr = retstr & "8"c
				Case 9
					retstr = retstr & "9"c
				Case 10
					retstr = retstr & "A"c
				Case 11
					retstr = retstr & "B"c
				Case 12
					retstr = retstr & "C"c
				Case 13
					retstr = retstr & "D"c
				Case 14
					retstr = retstr & "E"c
				Case 15
					retstr = retstr & "F"c
			End Select
			i = Dval And &HF
			Select Case i
				Case 0
					retstr = retstr & "0"c
				Case 1
					retstr = retstr & "1"c
				Case 2
					retstr = retstr & "2"c
				Case 3
					retstr = retstr & "3"c
				Case 4
					retstr = retstr & "4"c
				Case 5
					retstr = retstr & "5"c
				Case 6
					retstr = retstr & "6"c
				Case 7
					retstr = retstr & "7"c
				Case 8
					retstr = retstr & "8"c
				Case 9
					retstr = retstr & "9"c
				Case 10
					retstr = retstr & "A"c
				Case 11
					retstr = retstr & "B"c
				Case 12
					retstr = retstr & "C"c
				Case 13
					retstr = retstr & "D"c
				Case 14
					retstr = retstr & "E"c
				Case 15
					retstr = retstr & "F"c
			End Select
			result = retstr
			Return result
		End Function

		Public Shared Function OpenPort(ByVal PortName As String) As Boolean
			Dim result As Boolean
			Dim res As Integer
			Dim NoOfDevs As Integer
			Dim i As Integer
			Dim J As Integer
			Dim Name As String
			Dim DualName As String
			Dim done As Boolean
			PortAIsOpen = False
			result = False
			Name = ""
			DualName = PortName
            res = GetFTDeviceCount()
            If res <> FT_OK Then
                Return result
            End If
            NoOfDevs = FT_Device_Count
            J = 0
            If NoOfDevs > 0 Then
                Do
                    Do
                        res = GetFTDeviceDescription(J)
                        If (res <> FT_OK) Then
                            J = J + 1
                        End If
                    Loop While Not ((res = FT_OK) OrElse (J = NoOfDevs))
                    If res <> FT_OK Then
                        Return result
                    End If
                    done = False
                    i = 1
                    ' Name := '';
                    ' repeat
                    ' if ORD(FT_Device_String_Buffer[i]) <> 0 then
                    ' begin
                    Name = FT_Description
                    ' Name := Name + FT_Device_String_Buffer[i];
                    ' end
                    ' else
                    ' begin
                    ' done := true;
                    ' end;
                    ' i := i + 1;
                    ' until done;
                    J = J + 1
                Loop While Not ((J = NoOfDevs) OrElse (Name = DualName))
            End If
            If (Name = DualName) Then
                res = Open_USB_Device_By_Description(Name)
                If res <> FT_OK Then
                    Return result
                End If
                result = True
                res = Get_USB_Device_Queue_Status()
                If res <> FT_OK Then
                    Return result
                End If
                PortAIsOpen = True
            Else
                result = False
            End If
            Return result
        End Function

        Public Shared Sub List_Devs(ByRef My_Names_Ptr() As String)
            Dim res As Integer
            Dim NoOfDevs As Integer
            Dim J As Integer
            Dim k As Integer
            Dim Name As String
            PortAIsOpen = False
            Name = ""
            res = GetFTDeviceCount()
            If res <> FT_OK Then
                Return
            End If
            NoOfDevs = FT_Device_Count
            J = 0
            k = 1
            If NoOfDevs > 0 Then
                Do
                    res = GetFTDeviceDescription(J)
                    If res = FT_OK Then
                        My_Names_Ptr(k) = FT_Description
                        k = k + 1
                    End If
                    J = J + 1
                Loop While Not ((J = NoOfDevs))
            End If
        End Sub

        Public Shared Sub ClosePort()
            Dim res As Integer
            If PortAIsOpen Then
                res = Close_USB_Device()
            End If
            PortAIsOpen = False
        End Sub

        Public Shared Function Init_Controller(ByVal DName As String) As Boolean
            Dim result As Boolean
            Dim passed As Boolean
            Dim res As Integer
            result = False
            passed = OpenPort(DName)
            If passed Then
                res = Set_USB_Device_Latency_Timer() '(16)
                If (res = FT_OK) Then
                    result = True
                End If
            End If
            Return result
        End Function

        Public Shared Function write_value(ByVal enable As Byte, ByVal data As Byte, ByVal DName As String) As Boolean
            Dim result As Boolean
            Dim passed As Boolean
            Dim tmpval As Byte
            Dim res As Integer
            result = False
            passed = Init_Controller(DName)
            If passed Then
                tmpval = CByte(data And &HF)
                tmpval = CByte(tmpval Or ((enable And &HF) * 16))
                res = Set_USB_Device_Bit_Mode(tmpval, &H20)
                ' enable CBitBang
                If (res = FT_OK) Then
                    result = True
                End If
                ClosePort()
            End If
            Return result
        End Function

        Public Shared Function read_value(ByRef data As Byte, ByVal DName As String) As Boolean
            Dim result As Boolean
            Dim passed As Boolean
            Dim res As Integer
            result = False
            passed = Init_Controller(DName)
            If passed Then
                res = Get_USB_Device_Bit_Mode(data)
                If (res = FT_OK) Then
                    result = True
                End If
                ClosePort()
            End If
            Return result
        End Function

    End Class ' end iotst

End Namespace

