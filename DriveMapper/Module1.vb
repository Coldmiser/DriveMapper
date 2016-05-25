Module Module1
    Sub Main()
        Console.Write("Please enter your password: ")
        Dim password As String = String.Empty
        Dim info As ConsoleKeyInfo
        Do
            info = Console.ReadKey(True)
            If info.Key = ConsoleKey.Enter Then
                Exit Do
            Else
                password &= info.KeyChar
                Console.Write("*"c)
            End If
        Loop

        Console.WriteLine()
        'Console.WriteLine("Password is: " & password)
        MapDrive("E", "\\COMPUTER\e$", "COMPUTER\USER", "Password2")
        Console.Write(".")
        MapDrive("H", "\\cifs.company.com\users$\USER2", "company\USER2", password)
        Console.Write(".")
        MapDrive("N", "\\company.com\data\Enterprise\Temp", "company\USER2", password)
        Console.Write(".")
        MapDrive("S", "\\company.com\app\IT", "company\USER2", password)
        Console.Write(".")
        MapDrive("T", "\\company.com\data\IT", "company\USER2", password)
        Console.Write(".")
        MapDrive("U", "\\ppgulc1000\Scripting", "company\USER2", password)
        Console.Write(".")
        MapDrive("V", "\\company.com\app", "company\USER2", password)
        Console.Write(".")
        'Console.ReadKey()
    End Sub

    Public Declare Function WNetAddConnection2 Lib "mpr.dll" Alias "WNetAddConnection2A" (ByRef lpNetResource As NETRESOURCE, ByVal lpPassword As String, ByVal lpUserName As String, ByVal dwFlags As Integer) As Integer
    Public Declare Function WNetCancelConnection2 Lib "mpr" Alias "WNetCancelConnection2A" (ByVal lpName As String, ByVal dwFlags As Integer, ByVal fForce As Integer) As Integer
    Public Const ForceDisconnect As Integer = 1
    Public Const RESOURCETYPE_DISK As Long = &H1
    Private Const ERROR_BAD_NETPATH = 53&
    Private Const ERROR_NETWORK_ACCESS_DENIED = 65&
    Private Const ERROR_INVALID_PASSWORD = 86&
    Private Const ERROR_NETWORK_BUSY = 54&

    Public Structure NETRESOURCE
        Public dwScope As Integer
        Public dwType As Integer
        Public dwDisplayType As Integer
        Public dwUsage As Integer
        Public lpLocalName As String
        Public lpRemoteName As String
        Public lpComment As String
        Public lpProvider As String
    End Structure

    Public Function MapDrive(ByVal DriveLetter As String, ByVal UNCPath As String, ByVal strUsername As String, ByVal strPassword As String) As Boolean
        Dim nr As NETRESOURCE

        nr = New NETRESOURCE
        nr.lpRemoteName = UNCPath
        nr.lpLocalName = DriveLetter & ":"
        nr.lpProvider = Nothing
        nr.dwType = RESOURCETYPE_DISK

        Dim result As Integer
        result = WNetAddConnection2(nr, strPassword, strUsername, 0)

        If result = 0 Then
            Return True
        Else
            Select Case result
                Case ERROR_BAD_NETPATH
                    Console.WriteLine("QA4001I", "Bad path could not connect to Directory")
                Case ERROR_INVALID_PASSWORD
                    Console.WriteLine("QA4002I", "Invalid password could not connect to Directory")
                Case ERROR_NETWORK_ACCESS_DENIED
                    Console.WriteLine("QA4003I", "Network access denied could not connect to Directory")
                Case ERROR_NETWORK_BUSY
                    Console.WriteLine("QA4004I", "Network busy could not connect to Directory")
            End Select
            Return False
        End If

    End Function
End Module
