Imports System
Imports System.IO
Imports System.Security
Imports System.Security.Cryptography
Imports System.Runtime.InteropServices
Imports System.Text

Module Module1
    Dim EncryptionKey As String = "N0w 1s 4he 41m3 f0r a!! g00d m3n t0 c0m3 t0 4h3 a1d 0f 4h31r c0untry"
    Sub Main()
        Dim FullFileName = System.Reflection.Assembly.GetExecutingAssembly().Location
        Dim Path = System.IO.Path.GetDirectoryName(FullFileName)
        'Dim FileName = System.IO.Path.GetFileName(FullFileName))
        Dim ShortFileName = System.IO.Path.GetFileNameWithoutExtension(FullFileName)
        Dim DATFile = Path + "\" + ShortFileName + ".dat"
        If My.Application.CommandLineArgs.Count > 0 Then
            If My.Application.CommandLineArgs.Count = 4 Then
                'Here is where you put your information on building the INI file
                'Console.WriteLine(If(File.Exists(DATFile), "File exists.", "File does not exist."))
                Dim strArg() As String
                strArg = Command().Split(" ")
                Dim WriteReturn = (WriteINI(DATFile, strArg(0), strArg(1), strArg(2), strArg(3)))
                'Console.WriteLine(String.Join(Environment.NewLine, WriteReturn))
                Console.WriteLine("Drive Mapping Added Successfully")
            Else
                'Usage Statement (How can you write pages of stuff quickly?
                Console.WriteLine("Usage:")
                Console.WriteLine("")
                Console.WriteLine(ShortFileName)
                'Console.WriteLine(FileName)
            End If
        Else
            If (File.Exists(DATFile)) Then
                Dim ReadReturn = (ReadINI(DATFile))
                'Console.WriteLine("out")
                'Console.WriteLine(String.Join(Environment.NewLine, ReadReturn))
            Else
                'OK, let's map some network drives
                Console.Write("Please enter your password: ")
                Dim password As String = String.Empty
                Dim info As ConsoleKeyInfo
                Do
                    info = Console.ReadKey(True)
                    If info.Key = ConsoleKey.Enter Then
                        Exit Do
                    Else
                        'Replace the key pressed with an asterisk
                        password &= info.KeyChar
                        Console.Write("*"c)
                    End If
                Loop

                Console.WriteLine()
                'Console.WriteLine("Password is: " & password)
                'MapDrive("E", "\\COMPUTER\e$", "COMPUTER\USER", "Password2")
                'Console.Write("E")
                'MapDrive("H", "\\cifs.company.com\users$\USER2", "company\USER2", password)
                'Console.Write("H")
                'MapDrive("N", "\\company.com\data\Enterprise\Temp", "company\USER2", password)
                'Console.Write("N")
                'MapDrive("P", "\\company.com\app\IT", "company\USER2", password)
                'Console.Write("P")
                'MapDrive("S", "\\company.com\app\IT", "company\USER2", password)
                'Console.Write("S")
                'MapDrive("T", "\\company.com\data\IT", "company\USER2", password)
                'Console.Write("T")
                'MapDrive("U", "\\ppgulc1000\Scripting", "company\USER2", password)
                'Console.Write("U")
                'MapDrive("V", "\\company.com\app", "company\USER2", password)
                'Console.Write("V")
                'Console.ReadKey()
            End If

        End If
    End Sub

    Public Declare Function WNetAddConnection2 Lib "mpr.dll" Alias "WNetAddConnection2A" (ByRef lpNetResource As NETRESOURCE, ByVal lpPassword As String, ByVal lpUserName As String, ByVal dwFlags As Integer) As Integer
    Public Declare Function WNetCancelConnection2 Lib "mpr" Alias "WNetCancelConnection2A" (ByVal lpName As String, ByVal dwFlags As Integer, ByVal fForce As Integer) As Integer
    Public Const ForceDisconnect As Integer = 1
    Public Const RESOURCETYPE_DISK As Long = &H1
    'System Codes :  https://msdn.microsoft.com/en-us/library/windows/desktop/ms681381(v=vs.85).aspx
    'Network Codes:  https://msdn.microsoft.com/en-us/library/windows/desktop/aa385413(v=vs.85).aspx
    Private Const ERROR_BAD_NETPATH = 53&
    Private Const ERROR_NETWORK_BUSY = 54&
    Private Const ERROR_NETWORK_ACCESS_DENIED = 65&
    Private Const ERROR_BAD_NET_NAME = 67&
    Private Const ERROR_ALREADY_ASSIGNED = 85&
    Private Const ERROR_INVALID_PASSWORD = 86&
    Public NotInheritable Class Simple3Des
        Private TripleDes As New TripleDESCryptoServiceProvider

        Private Function TruncateHash(
        ByVal key As String,
        ByVal length As Integer) As Byte()

            Dim sha1 As New SHA1CryptoServiceProvider

            ' Hash the key.
            Dim keyBytes() As Byte =
            System.Text.Encoding.Unicode.GetBytes(key)
            Dim hash() As Byte = sha1.ComputeHash(keyBytes)

            ' Truncate or pad the hash.
            ReDim Preserve hash(length - 1)
            Return hash
        End Function

        Sub New(ByVal key As String)
            ' Initialize the crypto provider.
            TripleDes.Key = TruncateHash(key, TripleDes.KeySize \ 8)
            TripleDes.IV = TruncateHash("", TripleDes.BlockSize \ 8)
        End Sub
        Public Function EncryptData(
    ByVal plaintext As String) As String

            ' Convert the plaintext string to a byte array.
            Dim plaintextBytes() As Byte =
        System.Text.Encoding.Unicode.GetBytes(plaintext)

            ' Create the stream.
            Dim ms As New System.IO.MemoryStream
            ' Create the encoder to write to the stream.
            Dim encStream As New CryptoStream(ms,
        TripleDes.CreateEncryptor(),
        System.Security.Cryptography.CryptoStreamMode.Write)

            ' Use the crypto stream to write the byte array to the stream.
            encStream.Write(plaintextBytes, 0, plaintextBytes.Length)
            encStream.FlushFinalBlock()

            ' Convert the encrypted stream to a printable string.
            Return Convert.ToBase64String(ms.ToArray)
        End Function

        Public Function DecryptData(
    ByVal encryptedtext As String) As String

            ' Convert the encrypted text string to a byte array.
            Dim encryptedBytes() As Byte = Convert.FromBase64String(encryptedtext)

            ' Create the stream.
            Dim ms As New System.IO.MemoryStream
            ' Create the decoder to write to the stream.
            Dim decStream As New CryptoStream(ms,
        TripleDes.CreateDecryptor(),
        System.Security.Cryptography.CryptoStreamMode.Write)

            ' Use the crypto stream to write the byte array to the stream.
            decStream.Write(encryptedBytes, 0, encryptedBytes.Length)
            decStream.FlushFinalBlock()

            ' Convert the plaintext stream to a string.
            Return System.Text.Encoding.Unicode.GetString(ms.ToArray)
        End Function
    End Class


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

    Public Function MapDrive(ByVal DriveLetter As String,
                             ByVal UNCPath As String,
                             ByVal strUsername As String,
                             ByVal strPassword As String) As Boolean

        Dim nr As NETRESOURCE

        nr = New NETRESOURCE
        nr.lpRemoteName = UNCPath
        nr.lpLocalName = DriveLetter & ":"
        nr.lpProvider = Nothing
        nr.dwType = RESOURCETYPE_DISK

        Dim result As Integer
        result = WNetAddConnection2(nr, strPassword, strUsername, 0)

        If result = 0 Then
            Console.ForegroundColor = ConsoleColor.White
            Console.BackgroundColor = ConsoleColor.DarkGreen
            Console.Write(DriveLetter)
            Console.ResetColor()
            Return True
        Else
            Console.ForegroundColor = ConsoleColor.Black
            Console.BackgroundColor = ConsoleColor.Red
            Select Case result
                Case ERROR_BAD_NETPATH
                    '53 (0x35) The network path was Not found.
                    'Console.WriteLine("QA4001I", "Bad path could not connect to Directory")
                    Console.WriteLine(DriveLetter + ":  Bad path could not connect to Directory")
                Case ERROR_NETWORK_BUSY
                    '54 (0x36) The network Is busy.
                    'Console.WriteLine("QA4004I", "Network busy could not connect to Directory")
                    Console.WriteLine(DriveLetter + ":  Network busy could not connect to Directory")
                Case ERROR_NETWORK_ACCESS_DENIED
                    '65 (0x41) Network access Is denied.
                    'Console.WriteLine("QA4003I", "Network access denied could not connect to Directory")
                    Console.WriteLine(DriveLetter + ":  Network access denied could not connect to Directory")
                Case ERROR_BAD_NET_NAME
                    '67 (0x43) The network name cannot be found.
                    Console.WriteLine(DriveLetter + ":  The network name cannot be found")
                Case ERROR_ALREADY_ASSIGNED
                    '85 (0x55) The local device name Is already in use.
                    Console.WriteLine(DriveLetter + ":  The local device name Is already in use.")
                Case ERROR_INVALID_PASSWORD
                    '86 (0x56) The specified network password Is Not correct.
                    'Console.WriteLine("QA4002I", "Invalid password could not connect to Directory")
                    Console.WriteLine(DriveLetter + ":  Invalid password could not connect to Directory")
                Case Else
                    Console.WriteLine(DriveLetter + ":  " + result)
            End Select
            Console.ResetColor()
            Return False
        End If
    End Function

    Public Function ReadINI(ByVal DATFile As String) As String()
        Dim cleartext(3) As String
        Try
            ' Open the file using a stream reader.
            Dim list As New List(Of String)
            Using sr As New StreamReader(DATFile)
                Dim line As String
                ' Read first line.
                line = sr.ReadLine
                ' Loop over each line in file, While list is Not Nothing.
                Do While (Not line Is Nothing)
                    ' Add this line to list.
                    list.Add(line)
                    ' Display to console.
                    'Console.WriteLine(line)
                    ' Read in the next line.
                    Dim LineArray() As String = Split(line, vbTab)


                    Dim LastNonEmpty As Integer = -1
                    For i As Integer = 0 To LineArray.Length - 1
                        If LineArray(i) <> "" Then
                            LastNonEmpty += 1
                            LineArray(LastNonEmpty) = LineArray(i)
                        End If
                    Next

                    Dim wrapper As New Simple3Des(EncryptionKey)
                    cleartext(0) = wrapper.DecryptData(LineArray(0))
                    cleartext(1) = wrapper.DecryptData(LineArray(1))
                    cleartext(2) = wrapper.DecryptData(LineArray(2))
                    cleartext(3) = wrapper.DecryptData(LineArray(3))

                    Dim format As String = "{0,-2} {1,-40} {2,-20} {3,-10}"

                    ' Construct lines.
                    Dim line1 As String = String.Format(format, cleartext(0), cleartext(1), cleartext(2), "[password]")

                    ' Print them.
                    'Console.WriteLine(line1)
                    'Map That Drive!
                    MapDrive(cleartext(0), cleartext(1), cleartext(2), cleartext(3))


                    'Console.WriteLine("Decrypted:  " + cleartext(0) + vbTab + cleartext(1) + vbTab + cleartext(2) + vbTab + "[password]")
                    'Console.WriteLine(String.Join(vbTab, cleartext))

                    line = sr.ReadLine
                Loop
                Console.WriteLine("")

            End Using
        Catch e As Exception
            'Console.WriteLine("The file could not be read:")
            'Console.WriteLine(e.Message)
        End Try

        'Console.WriteLine("ReadINI " + DriveLetter)
        'Console.WriteLine("ReadINI " + UNCPath)
        'Console.WriteLine("ReadINI " + strUsername)
        'Console.WriteLine("ReadINI " + strPassword)
        Return cleartext
    End Function
    Public Function WriteINI(ByVal DATFile As String,
                             ByVal DriveLetter As String,
                             ByVal UNCPath As String,
                             ByVal strUsername As String,
                             ByVal strPassword As String) As String()

        'Console.WriteLine("WriteINI " + DriveLetter)
        'Console.WriteLine("WriteINI " + UNCPath)
        'Console.WriteLine("WriteINI " + strUsername)
        'Console.WriteLine("WriteINI " + strPassword)


        Dim wrapper As New Simple3Des(EncryptionKey)
        Dim cipher(3) As String
        cipher(0) = wrapper.EncryptData(DriveLetter)
        cipher(1) = wrapper.EncryptData(UNCPath)
        cipher(2) = wrapper.EncryptData(strUsername)
        cipher(3) = wrapper.EncryptData(strPassword)

        Dim AllCreds = cipher(0) + vbTab + cipher(1) + vbTab + cipher(2) + vbTab + cipher(3)
        'Console.WriteLine("Encrypted data: " + AllCreds)
        If Not File.Exists(DATFile) Then
            'Create a file to write to
            Using sw As StreamWriter = File.CreateText(DATFile)
                sw.WriteLine(AllCreds)
            End Using
        Else
            'Or append to the current one
            Using sw As StreamWriter = File.AppendText(DATFile)
                sw.WriteLine(AllCreds)
            End Using
        End If

        Return cipher
    End Function

End Module
