'Option Strict On
Imports HoMIDom
Imports HoMIDom.HoMIDom
Imports HoMIDom.HoMIDom.Server
Imports HoMIDom.HoMIDom.Device
Imports STRGS = Microsoft.VisualBasic.Strings
Imports System.IO.Ports
Imports VB = Microsoft.VisualBasic

' Auteur : HoMIDoM
' Date : 03/03/2013

''' <summary>Class Driver_MemoryLink, permet de commander et recevoir des ordres avec les périphériques supportant le protocole MemoryLink</summary>
''' 
<Serializable()> Public Class Driver_MemoryLink
    Implements HoMIDom.HoMIDom.IDriver

#Region "Variables génériques"
    '!!!Attention les variables ci-dessous doivent avoir une valeur par défaut obligatoirement
    'aller sur l'adresse http://www.somacon.com/p113.php pour avoir un ID
    Dim _ID As String = "7B5F099E-85A0-11E2-81FC-248E6188709A"
    Dim _Nom As String = "MemoryLink"
    Dim _Enable As Boolean = False
    Dim _Description As String = "MemoryLink"
    Dim _StartAuto As Boolean = False
    Dim _Protocol As String = "COM"
    Dim _IsConnect As Boolean = False
    Dim _IP_TCP As String = "@"
    Dim _Port_TCP As String = "@"
    Dim _IP_UDP As String = "@"
    Dim _Port_UDP As String = "@"
    Dim _Com As String = "COM1"
    Dim _Refresh As Integer = 0
    Dim _Modele As String = "Omron"
    Dim _Version As String = My.Application.Info.Version.ToString
    Dim _OsPlatform As String = "3264"
    Dim _Picture As String = ""
    Dim _Server As HoMIDom.HoMIDom.Server
    Dim _Device As HoMIDom.HoMIDom.Device
    Dim _DeviceSupport As New ArrayList
    Dim _Parametres As New ArrayList
    Dim _LabelsDriver As New ArrayList
    Dim _LabelsDevice As New ArrayList
    Dim WithEvents MyTimer As New Timers.Timer
    Dim _IdSrv As String
    Dim _DeviceCommandPlus As New List(Of HoMIDom.HoMIDom.Device.DeviceCommande)
    Dim _AutoDiscover As Boolean = False
    Dim _DEBUG As Boolean = False

#End Region

#Region "Variables internes"

    Private BufferIn(8192) As Byte	
    
    Dim str_to_int As New Dictionary(Of String, Integer)

	Dim TabCom As String() = {"COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9", "NOTHING"}

    Private m_port_name As String
    Private m_DebutTrame As Boolean = False
    Private m_bytecnt As Integer = 0
    Private m_messcnt As Integer = 0
    Private cnt As Integer = 0
    Private msg1 As String = ""

    Private adresse As UInt16
    Private data1(100) As UInt16
    Private longueur As UInt16
    Private commande As String



    Private m_recbuf(200) As Byte

    Private m_trame As Boolean = False
    Private m_IsConnectPort As Boolean = False
	
    Private WithEvents m_SerialPort As New System.IO.Ports.SerialPort

    Public Property port_name() As String
        Get
            Return m_port_name
        End Get
        Set(ByVal value As String)
            m_port_name = value
        End Set
    End Property

    Public Property DebutTrame() As Boolean
        Get
            Return m_DebutTrame
        End Get
        Set(ByVal value As Boolean)
            m_DebutTrame = value
        End Set
    End Property

    Public Property bytecnt() As Integer
        Get
            Return m_bytecnt
        End Get
        Set(ByVal value As Integer)
            m_bytecnt = value
        End Set
    End Property

    Public Property messcnt() As Integer
        Get
            Return m_messcnt
        End Get
        Set(ByVal value As Integer)
            m_messcnt = value
        End Set

    End Property

    Public Property recbuf As Byte()
        Get
            Return m_recbuf
        End Get
        Set(ByVal value As Byte())
            m_recbuf = value
        End Set
    End Property

    Public Property trame() As Boolean
        Get
            Return m_trame
        End Get
        Set(ByVal value As Boolean)
            m_trame = value
        End Set
    End Property

    Public Property IsConnectPort() As Boolean
        Get
            Return m_IsConnectPort
        End Get
        Set(ByVal value As Boolean)
            m_IsConnectPort = value
        End Set
    End Property

    Public Property SerialPort() As System.IO.Ports.SerialPort
        Get
            Return m_SerialPort
        End Get
        Set(ByVal value As System.IO.Ports.SerialPort)
            m_SerialPort = value
        End Set

    End Property

#End Region

#Region "Propriétés génériques"

    Public Event DriverEvent(ByVal DriveName As String, ByVal TypeEvent As String, ByVal Parametre As Object) Implements HoMIDom.HoMIDom.IDriver.DriverEvent

    Public WriteOnly Property IdSrv As String Implements HoMIDom.HoMIDom.IDriver.IdSrv
        Set(ByVal value As String)
            _IdSrv = value
        End Set
    End Property

    Public Property COM() As String Implements HoMIDom.HoMIDom.IDriver.COM
        Get
            Return _Com
        End Get
        Set(ByVal value As String)
            _Com = value
        End Set
    End Property
    Public ReadOnly Property Description() As String Implements HoMIDom.HoMIDom.IDriver.Description
        Get
            Return _Description
        End Get
    End Property
    Public ReadOnly Property DeviceSupport() As System.Collections.ArrayList Implements HoMIDom.HoMIDom.IDriver.DeviceSupport
        Get
            Return _DeviceSupport
        End Get
    End Property
    Public Property Parametres() As System.Collections.ArrayList Implements HoMIDom.HoMIDom.IDriver.Parametres
        Get
            Return _Parametres
        End Get
        Set(ByVal value As System.Collections.ArrayList)
            _Parametres = value
        End Set
    End Property

    Public Property LabelsDriver() As System.Collections.ArrayList Implements HoMIDom.HoMIDom.IDriver.LabelsDriver
        Get
            Return _LabelsDriver
        End Get
        Set(ByVal value As System.Collections.ArrayList)
            _LabelsDriver = value
        End Set
    End Property
    Public Property LabelsDevice() As System.Collections.ArrayList Implements HoMIDom.HoMIDom.IDriver.LabelsDevice
        Get
            Return _LabelsDevice
        End Get
        Set(ByVal value As System.Collections.ArrayList)
            _LabelsDevice = value
        End Set
    End Property



    Public Property Enable() As Boolean Implements HoMIDom.HoMIDom.IDriver.Enable
        Get
            Return _Enable
        End Get
        Set(ByVal value As Boolean)
            _Enable = value
        End Set
    End Property
    Public ReadOnly Property ID() As String Implements HoMIDom.HoMIDom.IDriver.ID
        Get
            Return _ID
        End Get
    End Property
    Public Property IP_TCP() As String Implements HoMIDom.HoMIDom.IDriver.IP_TCP
        Get
            Return _IP_TCP
        End Get
        Set(ByVal value As String)
            _IP_TCP = value
        End Set
    End Property
    Public Property IP_UDP() As String Implements HoMIDom.HoMIDom.IDriver.IP_UDP
        Get
            Return _IP_UDP
        End Get
        Set(ByVal value As String)
            _IP_UDP = value
        End Set
    End Property
    Public ReadOnly Property IsConnect() As Boolean Implements HoMIDom.HoMIDom.IDriver.IsConnect
        Get
            Return _IsConnect
        End Get
    End Property
    Public Property Modele() As String Implements HoMIDom.HoMIDom.IDriver.Modele
        Get
            Return _Modele
        End Get
        Set(ByVal value As String)
            _Modele = value
        End Set
    End Property
    Public ReadOnly Property Nom() As String Implements HoMIDom.HoMIDom.IDriver.Nom
        Get
            Return _Nom
        End Get
    End Property
    Public Property Picture() As String Implements HoMIDom.HoMIDom.IDriver.Picture
        Get
            Return _Picture
        End Get
        Set(ByVal value As String)
            _Picture = value
        End Set
    End Property
    Public Property Port_TCP() As String Implements HoMIDom.HoMIDom.IDriver.Port_TCP
        Get
            Return _Port_TCP
        End Get
        Set(ByVal value As String)
            _Port_TCP = value
        End Set
    End Property
    Public Property Port_UDP() As String Implements HoMIDom.HoMIDom.IDriver.Port_UDP
        Get
            Return _Port_UDP
        End Get
        Set(ByVal value As String)
            _Port_UDP = value
        End Set
    End Property
    Public ReadOnly Property Protocol() As String Implements HoMIDom.HoMIDom.IDriver.Protocol
        Get
            Return _Protocol
        End Get
    End Property
    Public Property Refresh() As Integer Implements HoMIDom.HoMIDom.IDriver.Refresh
        Get
            Return _Refresh
        End Get
        Set(ByVal value As Integer)
            _Refresh = value
        End Set
    End Property
    Public Property Server() As HoMIDom.HoMIDom.Server Implements HoMIDom.HoMIDom.IDriver.Server
        Get
            Return _Server
        End Get
        Set(ByVal value As HoMIDom.HoMIDom.Server)
            _Server = value
        End Set
    End Property
    Public ReadOnly Property Version() As String Implements HoMIDom.HoMIDom.IDriver.Version
        Get
            Return _Version
        End Get
    End Property
    Public ReadOnly Property OsPlatform() As String Implements HoMIDom.HoMIDom.IDriver.OsPlatform
        Get
            Return _OsPlatform
        End Get
    End Property
    Public Property StartAuto() As Boolean Implements HoMIDom.HoMIDom.IDriver.StartAuto
        Get
            Return _StartAuto
        End Get
        Set(ByVal value As Boolean)
            _StartAuto = value
        End Set
    End Property
    Public Property AutoDiscover() As Boolean Implements HoMIDom.HoMIDom.IDriver.AutoDiscover
        Get
            Return _AutoDiscover
        End Get
        Set(ByVal value As Boolean)
            _AutoDiscover = value
        End Set
    End Property
#End Region

#Region "Fonctions génériques"

    ''' <summary>
    ''' Retourne la liste des Commandes avancées
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetCommandPlus() As List(Of DeviceCommande)
        Return _DeviceCommandPlus
    End Function

    ''' <summary>Execute une commande avancée</summary>
    ''' <param name="MyDevice">Objet représentant le Device </param>
    ''' <param name="Command">Nom de la commande avancée à éxécuter</param>
    ''' <param name="Param">tableau de paramétres</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ExecuteCommand(ByVal MyDevice As Object, ByVal Command As String, Optional ByVal Param() As Object = Nothing) As Boolean
        Dim retour As Boolean = False
        Try
            If MyDevice IsNot Nothing Then
                'Pas de commande demandée donc erreur
                If Command = "" Then
                    Return False
                Else
                    Write(MyDevice, Command, Param(0), Param(1))
                    'Select Case UCase(Command)
                    '    Case ""
                    '    Case Else
                    'End Select
                    Return True
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            _Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, Me.Nom & " ExecuteCommand", "exception : " & ex.Message)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Permet de vérifier si un champ est valide
    ''' </summary>
    ''' <param name="Champ"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function VerifChamp(ByVal Champ As String, ByVal Value As Object) As String Implements HoMIDom.HoMIDom.IDriver.VerifChamp
        Try
            Dim retour As String = "0"
            Select Case UCase(Champ)


            End Select
            Return retour
        Catch ex As Exception
            Return "Une erreur est apparue lors de la vérification du champ " & Champ & ": " & ex.ToString
        End Try
    End Function

    ''' <summary>Démarrer le du driver</summary>
    ''' <remarks></remarks>
    Public Sub Start() Implements HoMIDom.HoMIDom.IDriver.Start
       
        'ouverture du port suivant le Port Com ou IP

        Dim retour As String = ""
        Dim retour2 As String = ""

        'récupération des paramétres avancés
        Try
            _DEBUG = _Parametres.Item(0).Valeur

        Catch ex As Exception
            _Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, "MemoryLink Start", "Erreur dans les paramétres avancés. utilisation des valeur par défaut" & ex.Message)
        End Try

        'ouverture du port suivant le Port Com
        Try
            If _Com <> "" Then
                port_name = _Com.ToUpper
                retour = ouvrir()
            Else
                retour = "ERR: Port Com non défini. Impossible d'ouvrir le port !"
            End If

            'traitement des messages de retour
            If (STRGS.Left(retour, 4) = "ERR:") Then
                IsConnectPort = False
                _IsConnect = False
                retour = STRGS.Right(retour, retour.Length - 5)
                _Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, "MemoryLink", "Driver non démarré : " & retour)
            Else
                AddHandler _Server.DeviceChanged, AddressOf DeviceChange
                _IsConnect = True
                IsConnectPort = True
                _Server.Log(TypeLog.INFO, TypeSource.DRIVER, "MemoryLink", retour)
            End If

        Catch ex As Exception
            _Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, "MemoryLink Start", ex.Message)
            _IsConnect = False
            IsConnectPort = False
        End Try
    End Sub

    ''' <summary>Arrêter le du driver</summary>
    ''' <remarks></remarks>
    Public Sub [Stop]() Implements HoMIDom.HoMIDom.IDriver.Stop
        Dim retour As String

        ' Recherche de tous les ports ouverts
        Try
            If IsConnectPort Then
                RemoveHandler _Server.DeviceChanged, AddressOf DeviceChange
                retour = fermer()
                _Server.Log(TypeLog.INFO, TypeSource.DRIVER, "MemoryLink Stop", retour)
            Else
                _Server.Log(TypeLog.INFO, TypeSource.DRIVER, "MemoryLink Stop", "Port " & _Com & " est déjà fermé")
            End If

        Catch ex As Exception
            _Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, "MemoryLink Stop", ex.Message)
        End Try
    End Sub

    ''' <summary>Re-Démarrer le du driver</summary>
    ''' <remarks></remarks>
    Public Sub Restart() Implements HoMIDom.HoMIDom.IDriver.Restart
        Try
            [Stop]()
            Start()
        Catch ex As Exception
            _Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, "MemoryLink ReStart", ex.Message)
        End Try
    End Sub

    ''' <summary>Intérroger un device</summary>
    ''' <param name="Objet">Objet représetant le device à interroger</param>
    ''' <remarks></remarks>
    Public Sub Read(ByVal Objet As Object) Implements HoMIDom.HoMIDom.IDriver.Read
        Try
            If _Enable = False Then
                _Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, Me.Nom & " Read", "Impossible d'effectuer un Read car le driver n'est pas Activé")
                Exit Sub
            End If
            If _IsConnect = False Then
                _Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, Me.Nom & " Read", "Impossible d'effectuer un Read car le driver n'est pas connecté à la carte")
                Exit Sub
            End If
            Dim _type As String = ""
            If CInt(Objet.adresse1) > 2999 Then
                _type = "Mots"
            Else
                _type = "Bits"
            End If
            ReadDM(_type, Objet.adresse1, 1)

        Catch ex As Exception
            _Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, Me.Nom & " Read", "Erreur : " & ex.ToString)
        End Try
    End Sub

    ''' <summary>Commander un device</summary>
    ''' <param name="Objet">Objet représetant le device à interroger</param>
    ''' <param name="Commande">La commande à passer</param>
    ''' <param name="Parametre1">La valeur à passer</param>
    ''' <param name="Parametre2"></param>
    ''' <remarks></remarks>
    Public Sub Write(ByVal Objet As Object, ByVal Commande As String, Optional ByVal Parametre1 As Object = Nothing, Optional ByVal Parametre2 As Object = Nothing) Implements HoMIDom.HoMIDom.IDriver.Write
        'Parametre1 = data1
        'Parametre2 = data2
        Dim sendtwice As Boolean = False
        Try
            If _Enable = False Then Exit Sub
            If _IsConnect = False Then
                _Server.Log(TypeLog.INFO, TypeSource.DRIVER, "MemoryLink Write", "Le driver n'est pas démarré, impossible d'écrire sur le port")
                Exit Sub
            End If
            If _DEBUG Then _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, "MemoryLink Write", "Ecriture de " & Objet.Name)
            If Parametre1 Is Nothing Then Parametre1 = 0
            If Parametre2 Is Nothing Then Parametre2 = 0

            If TypeOf Objet.Value Is Integer Or TypeOf Objet.Value Is Double Then
                ecrire(Objet.adresse1, Commande, Parametre1, Parametre2, sendtwice)
            End If

            If TypeOf Objet.Value Is Boolean And Not TypeOf Objet.Value Is Integer And Not TypeOf Objet.Value Is Double Then
                Parametre1 = str_to_int(Commande)
                ecrire(Objet.adresse1, Commande, Parametre1, Parametre2, sendtwice)
            End If

            Select Case Commande
                Case "ON", "OFF"
                    If Parametre1 = 0 Then traitement("OFF", Objet.adresse1, Commande, True) Else traitement("ON", Objet.adresse1, Commande, True)
                Case "DIM", "OUVERTURE"
                    traitement(CStr(Parametre1), Objet.adresse1, Commande, True)
            End Select

        Catch ex As Exception
            _Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, "MemoryLink Write", ex.Message)
        End Try
    End Sub

    ''' <summary>Fonction lancée lors de la suppression d'un device</summary>
    ''' <param name="DeviceId">Objet représetant le device à interroger</param>
    ''' <remarks></remarks>
    Public Sub DeleteDevice(ByVal DeviceId As String) Implements HoMIDom.HoMIDom.IDriver.DeleteDevice
        Try

        Catch ex As Exception
            _Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, "MemoryLink DeleteDevice", ex.Message)
        End Try
    End Sub

    ''' <summary>Fonction lancée lors de l'ajout d'un device</summary>
    ''' <param name="DeviceId">Objet représetant le device à interroger</param>
    ''' <remarks></remarks>
    Public Sub NewDevice(ByVal DeviceId As String) Implements HoMIDom.HoMIDom.IDriver.NewDevice
        Try

        Catch ex As Exception
            _Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, "MemoryLink NewDevice", ex.Message)
        End Try
    End Sub

    ''' <summary>ajout des commandes avancées pour les devices</summary>
    ''' <remarks></remarks>
    Private Sub add_devicecommande(ByVal nom As String, ByVal description As String, ByVal nbparam As Integer)
        Try
            Dim x As New DeviceCommande
            x.NameCommand = nom
            x.DescriptionCommand = description
            x.CountParam = nbparam
            _DeviceCommandPlus.Add(x)
        Catch ex As Exception
            _Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, "MemoryLink add_devicecommande", "Exception : " & ex.Message)
        End Try
    End Sub

    ''' <summary>ajout Libellé pour le Driver</summary>
    ''' <param name="nom">Nom du champ : HELP</param>
    ''' <param name="labelchamp">Nom à afficher : Aide</param>
    ''' <param name="tooltip">Tooltip à afficher au dessus du champs dans l'admin</param>
    ''' <remarks></remarks>
    Private Sub Add_LibelleDriver(ByVal Nom As String, ByVal Labelchamp As String, ByVal Tooltip As String, Optional ByVal Parametre As String = "")
        Try
            Dim y0 As New HoMIDom.HoMIDom.Driver.cLabels
            y0.LabelChamp = Labelchamp
            y0.NomChamp = UCase(Nom)
            y0.Tooltip = Tooltip
            y0.Parametre = Parametre
            _LabelsDriver.Add(y0)
        Catch ex As Exception
            _Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, Me.Nom & " add_devicecommande", "Exception : " & ex.Message)
        End Try
    End Sub

    ''' <summary>Ajout Libellé pour les Devices</summary>
    ''' <param name="nom">Nom du champ : HELP</param>
    ''' <param name="labelchamp">Nom à afficher : Aide, si = "@" alors le champ ne sera pas affiché</param>
    ''' <param name="tooltip">Tooltip à afficher au dessus du champs dans l'admin</param>
    ''' <remarks></remarks>
    Private Sub Add_LibelleDevice(ByVal Nom As String, ByVal Labelchamp As String, ByVal Tooltip As String, Optional ByVal Parametre As String = "")
        Try
            Dim ld0 As New HoMIDom.HoMIDom.Driver.cLabels
            ld0.LabelChamp = Labelchamp
            ld0.NomChamp = UCase(Nom)
            ld0.Tooltip = Tooltip
            ld0.Parametre = Parametre
            _LabelsDevice.Add(ld0)
        Catch ex As Exception
            _Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, Me.Nom & " add_devicecommande", "Exception : " & ex.Message)
        End Try
    End Sub

    ''' <summary>ajout de parametre avancés</summary>
    ''' <param name="nom">Nom du parametre (sans espace)</param>
    ''' <param name="description">Description du parametre</param>
    ''' <param name="valeur">Sa valeur</param>
    ''' <remarks></remarks>
    Private Sub add_paramavance(ByVal nom As String, ByVal description As String, ByVal valeur As Object)
        Try
            Dim x As New HoMIDom.HoMIDom.Driver.Parametre
            x.Nom = nom
            x.Description = description
            x.Valeur = valeur
            _Parametres.Add(x)
        Catch ex As Exception
            _Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, "MemoryLink add_devicecommande", "Exception : " & ex.Message)
        End Try
    End Sub

    ''' <summary>Creation d'un objet de type</summary>
    ''' <remarks></remarks>
    Public Sub New()
        Try
            _Version = Reflection.Assembly.GetExecutingAssembly.GetName.Version.ToString

            'Liste des devices compatibles
            _DeviceSupport.Add(ListeDevices.APPAREIL.ToString)
            _DeviceSupport.Add(ListeDevices.BAROMETRE.ToString)
            _DeviceSupport.Add(ListeDevices.BATTERIE.ToString)
            _DeviceSupport.Add(ListeDevices.COMPTEUR.ToString)
            _DeviceSupport.Add(ListeDevices.CONTACT.ToString)
            _DeviceSupport.Add(ListeDevices.DETECTEUR.ToString)
            _DeviceSupport.Add(ListeDevices.DIRECTIONVENT.ToString)
            _DeviceSupport.Add(ListeDevices.ENERGIEINSTANTANEE.ToString)
            _DeviceSupport.Add(ListeDevices.ENERGIETOTALE.ToString)
            _DeviceSupport.Add(ListeDevices.GENERIQUEBOOLEEN.ToString)
            _DeviceSupport.Add(ListeDevices.GENERIQUEVALUE.ToString)
            _DeviceSupport.Add(ListeDevices.HUMIDITE.ToString)
            _DeviceSupport.Add(ListeDevices.LAMPE.ToString)
            _DeviceSupport.Add(ListeDevices.PLUIECOURANT.ToString)
            _DeviceSupport.Add(ListeDevices.PLUIETOTAL.ToString)
            _DeviceSupport.Add(ListeDevices.SWITCH.ToString)
            _DeviceSupport.Add(ListeDevices.TELECOMMANDE.ToString)
            _DeviceSupport.Add(ListeDevices.TEMPERATURE.ToString)
            _DeviceSupport.Add(ListeDevices.TEMPERATURECONSIGNE.ToString)
            _DeviceSupport.Add(ListeDevices.UV.ToString)
            _DeviceSupport.Add(ListeDevices.VITESSEVENT.ToString)
            _DeviceSupport.Add(ListeDevices.VOLET.ToString)

            'Parametres avancés
            add_paramavance("Rafraichissement de lecture", "le temps en millisecondes entre les demandes de lecture", 2000)
            add_paramavance("Premier mot de lecture", "Adresse du premier mot à lire dans l'automate", 256)
            add_paramavance("Premier mot d'écriture", "Adresse du premier mot à écrire dans l'automate", 100)
            add_paramavance("Numéro Unit", "Numéro d'identification de l'unité a accéder", 0)
            add_paramavance("Debug", "Activer le Debug complet (True/False)", False)

            'ajout des commandes avancées pour les devices
            add_devicecommande("OFF", "Eteint tous les appareils du meme range que ce device", 0)
            add_devicecommande("ON", "Allume toutes les lampes du meme range que ce device", 0)
            add_devicecommande("DIM", "Variation, parametre = Variation", 1)

            'Libellé Driver
            Add_LibelleDriver("HELP", "Aide...", "Pas d'aide actuellement...")

            'Libellé Device
            Add_LibelleDevice("ADRESSE1", "Adresse", "Adresse du composant")
            Add_LibelleDevice("ADRESSE2", "@", "")
            Add_LibelleDevice("SOLO", "@", "")
            Add_LibelleDevice("REFRESH", "Rafraichissement", "Mettre 1 ou 0 pour que la valeur soit rafraichi ou pas")
            Add_LibelleDevice("LASTCHANGEDUREE", "@", "")

            'dictionnaire Commande STR -> INT
            str_to_int.Add("OFF", 0)
            str_to_int.Add("ON", 1)

        Catch ex As Exception
            _Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, "MemoryLink New", "Exception : " & ex.Message)
        End Try
    End Sub

    ''' <summary>Si refresh >0 gestion du timer</summary>
    ''' <remarks>PAS UTILISE CAR IL FAUT LANCER UN TIMER QUI LANCE/ARRETE CETTE FONCTION dans Start/Stop</remarks>
    Private Sub TimerTick(ByVal source As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles MyTimer.Elapsed
       
        If _Enable = False Then Exit Sub

        Try
            
        Catch ex As Exception
            _Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, "MemoryLink Read ", ex.Message)
        End Try

    End Sub

    Sub DeviceChange(ByVal DeviceID, ByVal valeurString)
        Try
            Dim genericDevice As templateDevice = _Server.ReturnDeviceById(_IdSrv, DeviceID)
            If genericDevice.DriverID = _ID Then
                If TypeOf genericDevice.Value Is Double Or TypeOf genericDevice.Value Is Integer Then
                    Write(genericDevice, "", CInt(valeurString))
                End If
            End If
        Catch ex As Exception
            _Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, "Memorylink Device Change", ex.Message)
        End Try
    End Sub

#End Region

#Region "Fonctions internes"

    ''' <summary>Ouvrir le port MemoryLink</summary>
    ''' <remarks></remarks>
    Private Function ouvrir() As String
        Try
            'ouverture du port
            If Not IsConnectPort And Not (SerialPort.IsOpen()) Then
                ' Test d'ouveture du port Com du controleur 
                SerialPort.PortName = port_name 'nom du port : COM1,COM2, COM3...   
                SerialPort.Open()
                ' Le port existe ==> le controleur est present
                If SerialPort.IsOpen() Then
                    SerialPort.Close()
                Else
                    Return ("ERR: Port " & port_name & " impossible à ouvrir")
                    Exit Function
                End If

                        SerialPort.BaudRate = 19200  'vitesse du port 300, 600, 1200, 2400, 9600, 14400, 19200, 38400, 57600, 115200
                        SerialPort.Parity = IO.Ports.Parity.None ' parité paire
                        SerialPort.StopBits = IO.Ports.StopBits.One 'un bit d'arrêt par octet
                        SerialPort.DataBits = 8 'nombre de bit par octet

                SerialPort.ReadTimeout = 3000
                SerialPort.WriteTimeout = 5000
                SerialPort.RtsEnable = True  'ligne Rts désactivé
                SerialPort.DtrEnable = True  'ligne Dtr désactivé
                SerialPort.Open()
                AddHandler SerialPort.DataReceived, New SerialDataReceivedEventHandler(AddressOf DataReceived)

                IsConnectPort = True
                _IsConnect = True

                Return ("Port " & port_name & " ouvert")
            Else
                Return ("Port " & port_name & " dejà ouvert")
            End If
        Catch ex As Exception
            Return ("ERR: " & ex.Message)
        End Try
    End Function

    ''' <summary>Fermer le port TeleInfo</summary>
    ''' <remarks></remarks>
    Private Function fermer() As String
        Try
            If IsConnectPort Then

                SerialPort.DiscardInBuffer()
                Threading.Thread.Sleep(1000)

                RemoveHandler SerialPort.DataReceived, AddressOf DataReceived

                If (Not (SerialPort Is Nothing)) Then ' The COM port exists.
                    If SerialPort.IsOpen Then
                        SerialPort.DiscardInBuffer()
                        SerialPort.Close()
                        'SerialPort.Dispose()
                        IsConnectPort = False

                        _IsConnect = False
                        Return ("Port " & port_name & " fermé")
                    Else
                        Return ("Port " & port_name & "  est déjà fermé")
                    End If
                Else
                    Return ("Port " & port_name & " n'existe pas")
                End If
            Else
                Return ("Port " & port_name & "  est déjà fermé (port_ouvert=false)")
            End If
        Catch ex As UnauthorizedAccessException
            Return ("ERR: Port " & port_name & " IGNORE")
            ' The port may have been removed. Ignore.
        End Try
        Return True
    End Function

    ''' <summary>Fonction lancée sur reception de données sur le port COM</summary>
    ''' <remarks></remarks>
    Private Sub DataReceived(ByVal sender As Object, ByVal e As SerialDataReceivedEventArgs)
        Try

            Dim BufferIn(8192) As Byte
            Dim count As Integer = 0

            ' Si le port est connecté - Reception des caracteres contenus dans le Buffer
            If IsConnectPort Then
                count = SerialPort.BytesToRead
                Try

                    If count Then
                        SerialPort.Read(BufferIn, 0, count)
                        
                        Dim msg As String = System.Text.Encoding.ASCII.GetString(BufferIn)

                        If _DEBUG Then _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " DataReceived : ", msg)
                        For i As Integer = 0 To count - 1
                            ProcessReceivedChar(BufferIn(i))
                        Next
                    Else
                        _Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, Me.Nom & " DataReceived", "Pas de donnée recue")
                    End If
                Catch ex As Exception
                    _Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, Me.Nom & " DataReceived ", "Exception : " & ex.Message)
                End Try
            Else
                If _DEBUG Then _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Port COM non connecté")
            End If

        Catch Ex As Exception
            _Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, Me.Nom & " Datareceived", "Exception : " & Ex.Message)
        End Try
    End Sub

    ''' <summary>Rassemble un message complet pour ensuite l'envoyer à displaymess</summary>
    ''' <param name="temp">Byte recu</param>
    ''' <remarks></remarks>
    Private Sub ProcessReceivedChar(ByVal temp As Byte)
        
        Try
           
            If (temp = CByte(27)) Then ' Debut de trame recu 
                DebutTrame = True
                bytecnt = 0
                messcnt = 0
                trame = False
                cnt = 0
                commande = ""
                adresse = 0
                longueur = 0
                data1(1) = 0
                msg1 = "Decodage:"
            ElseIf (DebutTrame And bytecnt = 2 And cnt > 0) Then
                commande = System.Text.Encoding.ASCII.GetString(recbuf, 0, 2)
                recbuf.Initialize()
                cnt = 0
                msg1 = msg1 & " commande = " & commande
            ElseIf (DebutTrame And bytecnt = 6 And cnt > 0) Then
                adresse = AsciiToWord(System.Text.Encoding.ASCII.GetString(recbuf, 0, 4))
                recbuf.Initialize()
                cnt = 0
                msg1 = msg1 & " et adresse = " & adresse
            ElseIf (DebutTrame And bytecnt = 8 And cnt > 0) Then
                longueur = CShort(System.Text.Encoding.ASCII.GetString(recbuf, 0, 2))
                recbuf.Initialize()
                cnt = 0
                msg1 = msg1 & " et longueur = " & longueur
            ElseIf (DebutTrame And bytecnt > 8 And cnt > 0 And temp = CByte(44) And messcnt + 1 < CInt(longueur)) Then ' debut d'info recu
                data1(messcnt + 1) = AsciiToWord(System.Text.Encoding.ASCII.GetString(recbuf, 0, cnt))
                messcnt += 1
                recbuf.Initialize()
                cnt = 0
                msg1 = msg1 & " et data" & messcnt & " = " & data1(messcnt)
            ElseIf (DebutTrame And temp = CByte(13)) Then 'Fin de trame recue
                data1(messcnt + 1) = AsciiToWord(System.Text.Encoding.ASCII.GetString(recbuf, 0, cnt - 2))
                recbuf.Initialize()
                cnt = 0
                messcnt += 1
                trame = True
                msg1 = msg1 & " et data finale = " & data1(messcnt) & " et bytecount = " & bytecnt & " et messcount = " & messcnt

            End If

            If temp <> CByte(27) And temp <> CByte(44) And temp <> CByte(13) Then
                recbuf.SetValue(temp, cnt)
                cnt += 1
                bytecnt += 1
            End If
            If temp = CByte(13) Then

                If _DEBUG Then _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " ProcessReceivedChar : ", msg1)

            End If

        Catch ex As Exception
            _Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, Me.Nom & " ProcessReceivedChar", ex.Message)
        End Try

        Try
            ' Une trame complete a été recue 
            Dim _adresse As String = "0"
            If trame Then
                If commande = "SM" Then
                    _adresse = CStr(CInt(adresse) + 3000)
                ElseIf commande = "SB" Then
                    _adresse = CStr(adresse)
                End If
                'Recherche si un device affecté
                Dim listedevices As New ArrayList
                listedevices = _Server.ReturnDeviceByAdresse1TypeDriver(_IdSrv, _adresse, "", Me._ID, True)

                If listedevices.Count > 0 Then

                    For Each j As Object In listedevices
                        If TypeOf j.Value Is Integer Or TypeOf j.Value Is Double Then
                            j.Value = data1(1)
                        End If
                        If TypeOf j.Value Is Boolean And Not TypeOf j.Value Is Integer And Not TypeOf j.Value Is Double And data1(1) > -1 And data1(1) < 2 Then
                            j.Value = CBool(data1(1))
                        End If
                    Next
                Else
                    _Server.Log(TypeLog.INFO, TypeSource.DRIVER, "Recherche composants", " Pas de composants trouvé à l'adresse " & _adresse)
                End If
            End If
        Catch ex As Exception
            _Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, Me.Nom & " Traitement Label - Traite Message", ex.Message)
        End Try
    End Sub

    ''' <summary>Ecrire sur le port MemoryLink</summary>
    ''' <param name="adresse">Adresse du device : A1...</param>
    ''' <param name="commande">commande à envoyer : ON, OFF...</param>
    ''' <param name="data1">voir description des actions plus haut ou doc MemoryLink</param>
    ''' <param name="data2">voir description des actions plus haut ou doc MemoryLink</param>
    ''' <param name="ecriretwice">Booleen : Ecrire l'ordre deux fois</param>
    ''' <remarks></remarks>

    Private Function ecrire(ByVal adresse As String, ByVal commande As String, Optional ByVal data1 As Integer = 0, Optional ByVal data2 As Integer = 0, Optional ByVal ecriretwice As Boolean = False) As String
        Dim _type As String = ""
        Dim _data As String = ""
        Dim _adresse As String = ""
        Try

            Dim _adresse2 As String = ""
            If _IsConnect Then

                Try

                    If CInt(adresse) > 3000 Then
                        _type = "Mots"
                        _data = data1
                        _adresse = CStr(CInt(adresse) - 3000)
                    Else
                        _type = "Bits"
                        _data = str_to_int(commande)
                        _adresse = adresse
                    End If

                    WriteDM(_type, _adresse, _data)

                Catch ex As Exception
                    Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, "MemoryLink", "MemoryLink Ecrire" & ex.Message)
                End Try

                'renvoie la valeur ecrite
                Return "VALUE"

            Else
                _Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, "MemoryLink Ecrire", "Port Fermé, impossible d ecrire : " & adresse & " : " & commande & " " & data1 & "-" & data2)
                Return ""
            End If
        Catch ex As Exception
            _Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, "MemoryLink Ecrire", "exception : " & ex.Message)
            Return "" 'on renvoi rien car il y a eu une erreur
        End Try
    End Function

    ''' <summary>Traite les paquets reçus</summary>
    ''' <remarks></remarks>
    Private Sub traitement(ByVal valeur As String, ByVal adresse As String, ByVal commande As String, ByVal erreursidevicepastrouve As Boolean)
        If valeur <> "" Then
            Try

                'Recherche si un device affecté
                Dim listedevices As New ArrayList
                listedevices = _Server.ReturnDeviceByAdresse1TypeDriver(_IdSrv, adresse, "", Me._ID, True)
                If IsNothing(listedevices) Then
                    _Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, "Process", "Communication impossible avec le serveur, l'IDsrv est peut être erroné : " & _IdSrv)
                    Exit Sub
                End If
                'un device trouvé on maj la value
                If (listedevices.Count = 1) Then
                    'correction valeur pour correspondre au type de value
                    If TypeOf listedevices.Item(0).Value Is Integer Then
                        If valeur = "ON" Then
                            listedevices.Item(0).Value = listedevices.Item(0).ValueMax
                        ElseIf valeur = "OFF" Then
                            listedevices.Item(0).Value = listedevices.Item(0).ValueMin
                        Else
                            listedevices.Item(0).Value = valeur
                        End If
                    ElseIf TypeOf listedevices.Item(0).Value Is Boolean Then
                        If valeur = "ON" Then
                            listedevices.Item(0).Value = True
                        ElseIf valeur = "OFF" Then
                            listedevices.Item(0).Value = False
                        Else
                            listedevices.Item(0).Value = True
                        End If
                    Else
                        listedevices.Item(0).Value = valeur
                    End If
                ElseIf (listedevices.Count > 1) Then
                    _Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, "Process", "Plusieurs devices correspondent à : " & adresse & ":" & valeur)
                Else
                    _Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, "Process", "Device non trouvé : " & adresse & ":" & valeur)

                    'Ajouter la gestion des composants bannis (si dans la liste des composant bannis alors on log en debug sinon onlog device non trouve empty)

                End If
            Catch ex As Exception
                _Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, "Traitement", "Exception : " & ex.Message & " --> " & adresse & " : " & commande & "-" & valeur)
            End Try
        End If
    End Sub

#End Region

#Region "MemoryLink"

    Private Function AsciiToWord(ByVal dataS As String) As UShort
        Try
            AsciiToWord = Convert.ToInt16(dataS, 16)
        Catch ex As Exception
            AsciiToWord = 0
            _Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, "MemoryLink exception", "AsciiToWord")
        End Try
    End Function
	
    Private Function WordToAscii(ByVal dataW As UShort, ByVal longueur As Integer) As String
        Try
            Dim zeroStr As String = ""
            For i As Integer = 1 To longueur
                zeroStr += "0"
            Next
            WordToAscii = VB.Right(zeroStr & CStr(dataW.ToString("X")), longueur)
        Catch ex As Exception
            WordToAscii = ""
            _Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, "MemoryLink exception", "WordToAscii")
        End Try
    End Function


    Public Function FCS(ByVal Command As String) As String

        Dim i As Short
        Dim Checksum As Short

        Try

            Checksum = 0

            For i = 1 To Len(Command)
                Checksum = Asc(Mid(Command, i, 1)) Xor Checksum
            Next i

            FCS = VB.Right("00" & Hex(Checksum), 2)

        Catch ex As Exception
            FCS = "00"
            _Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, "MemoryLink exception", "Checksum")
        End Try
    End Function


    Public Sub ReadDM(ByVal DMType As String, ByVal DMAddress As String, ByVal WordCount As Short)

        Dim Command As String = ""
        Dim Type As String = ""

        Try

            Select Case DMType
                Case "Mots"
                    Type = "RM"
                Case "Bits"
                    Type = "RB"
            End Select

            Command = Type & "0" & WordToAscii(DMAddress, 4) & VB.Right("00" & (Math.Truncate(WordCount)).ToString(), 2)
            'Command = Command & FCS(Command)

            If _DEBUG Then _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me._Nom & " ReadDM", ChrW(27) & Command & ChrW(13))

            SerialPort.Write(ChrW(27) & Command & ChrW(13))

        Catch ex As Exception
            _Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, "MemoryLink exception", "ReadDM")
        End Try

    End Sub

    Public Sub WriteDM(ByVal DMType As String, ByVal DMAddress As String, ByVal Data As String)

        Dim Command As String = ""
        Dim Type As String = ""
        Dim Data2 As String = ""

        Try

            Select Case DMType
                Case "Mots"
                    Type = "WM"
                    Data2 = VB.Right("0000" & Data, 4)
                Case "Bits"
                    Type = "WB"
                    Select Case Data
                        Case "0"
                            Data2 = "0"
                        Case "1"
                            Data2 = "8"
                    End Select
            End Select

            

            Command = Type & "0" & WordToAscii(DMAddress, 4) & "01" & Data2
            'Command = Command & FCS(Command)

            If _DEBUG Then _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me._Nom & " WriteDM", ChrW(27) & Command & ChrW(13))

            SerialPort.Write(ChrW(27) & Command & ChrW(13))

        Catch ex As Exception
            _Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, "MemoryLink exception", "WriteDM")
        End Try

    End Sub


#End Region

End Class
