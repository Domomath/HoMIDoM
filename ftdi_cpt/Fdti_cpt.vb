Public Class Ftdi_cpt

    Public Shared nom_interface As String = String.Empty
    Public Shared num_comport As Integer = 0
    Public Const nom_interface1 As String = "Interface USB 1 TIC"
    Public Const nom_interface2 As String = "Interface XBEE -> Compteur"
    Public Const nom_interface3 As String = "Interface USB -> Compteur"
    Public Event InfoReceived(ByVal data As String)

    Private Delegate Sub _Affiche_ASCII(ByVal donnee As String)

    Public ReadOnly Property PortCOM As Integer
        Get
            Return num_comport
        End Get
    End Property

    Public Sub cpt(ByVal num As Integer)
        Dim passed As Boolean
        passed = iotst.Units.iotst.write_value(num, num, nom_interface)
        ' fermeture des deux compteurs
        If passed Then
        Else
        End If
    End Sub

    ' procedure executée au lancement du programme
    Public Sub Start(ByVal nomInterface As String)

        Dim FT_result As Integer

        Try
            Search(nomInterface)

            nom_interface = nomInterface
            If FT_Handle > 0 Then
                FT_result = FT_OpenEx(nomInterface, FT_OPEN_BY_DESCRIPTION, FT_Handle)
            End If
            If FT_result = FT_OK Then
                FT_result = FT_GetComPortNumber(FT_Handle, num_comport)
                If FT_result <> FT_OK Then
                    RaiseEvent InfoReceived("Pas de port COM attribué !")
                End If
            Else
                RaiseEvent InfoReceived("Pas d'interface trouvée !")
            End If
            FT_Close(FT_Handle)

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Search(ByVal interface1 As String)

        Dim DeviceCount As Integer
        Dim TempDevString As String

        ' Get the number of device attached
        FT_Status = FT_GetNumberOfDevices(DeviceCount, vbNullChar, FT_LIST_NUMBER_ONLY)
        If FT_Status <> FT_OK Then
            Exit Sub
        End If

        ' Get description of device with index 0
        ' Allocate space for string variable
        For i = 0 To DeviceCount - 1

            ' Get serial number of device with index 0
            ' Allocate space for string variable
            TempDevString = Space(16)
            FT_Status = FT_GetDeviceString(i, TempDevString, FT_LIST_BY_INDEX Or FT_OPEN_BY_SERIAL_NUMBER)
            If FT_Status = FT_OK Then
                
                FT_Serial_Number = Microsoft.VisualBasic.Left(TempDevString, InStr(1, TempDevString, vbNullChar) - 1)

                TempDevString = Space(64)
                FT_Status = FT_GetDeviceString(i, TempDevString, FT_LIST_BY_INDEX Or FT_OPEN_BY_DESCRIPTION)
                If FT_Status <> FT_OK Then
                    Exit Sub
                End If
                FT_Description = Microsoft.VisualBasic.Left(TempDevString, InStr(1, TempDevString, vbNullChar) - 1)

            End If
            If FT_Description = interface1 Then Exit For
        Next
        'Open device by serial number
        FT_Status = FT_OpenByDescription(FT_Description, 2, FT_Handle)
        If FT_Status <> FT_OK Then
            RaiseEvent InfoReceived("Failed to open device.")
            Exit Sub
        End If


        ' Reset device
        FT_Status = FT_ResetDevice(FT_Handle)
        If FT_Status <> FT_OK Then
            Exit Sub
        End If

        ' Purge buffers
        FT_Status = FT_Purge(FT_Handle, FT_PURGE_RX Or FT_PURGE_TX)
        If FT_Status <> FT_OK Then
            Exit Sub
        End If

        ' Close device
        FT_Status = FT_Close(FT_Handle)
        If FT_Status <> FT_OK Then
            Exit Sub
        End If


    End Sub
End Class
