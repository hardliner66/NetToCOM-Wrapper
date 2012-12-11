Option Strict On
Imports Microsoft.VisualBasic
Imports System.Security.Cryptography
Imports System.IO
Imports System.Text

<ComClass(Crypto.ClassId, Crypto.InterfaceId, Crypto.EventsId)> _
Public Class Crypto
#Region "COM-GUIDs"
  ' Diese GUIDs stellen die COM-Identität für diese Klasse 
  ' und ihre COM-Schnittstellen bereit. Wenn Sie sie ändern, können vorhandene 
  ' Clients nicht mehr auf die Klasse zugreifen.
  Public Const ClassId As String = "18d6d903-dda1-45d5-9007-0567dc6b4f3b"
  Public Const InterfaceId As String = "fe7ce157-907a-41d6-8523-508e0dfb4e40"
  Public Const EventsId As String = "a6f657bd-ef40-4d79-a9a1-bc95ec78b634"
#End Region
  ' Eine erstellbare COM-Klasse muss eine Public Sub New() 
  ' ohne Parameter aufweisen. Andernfalls wird die Klasse 
  ' nicht in der COM-Registrierung registriert und kann nicht 
  ' über CreateObject erstellt werden.
  Public Sub New()
    MyBase.New()
  End Sub
  Public Function Encrypt(ByVal Text As String, ByVal Password As String) As String
    'Löst eine Exception aus wenn der eingegebene Text bzw das Passwort leer ist oder aus nur Whitespaces besteht
    If IsNullOrWhiteSpace(Text) Then
      If IsNothing(Text) Then
        Throw New ArgumentNullException("Text")
      Else
        Throw New ArgumentException("Argument cannot be empty or whitespace.", "Text")
      End If
      Exit Function
    End If
    If IsNullOrWhiteSpace(Password) Then
      If IsNothing(Password) Then
        Throw New ArgumentNullException("Password")
      Else
        Throw New ArgumentException("Argument cannot be empty or whitespace.", "Password")
      End If
      Exit Function
    End If

    Dim encdata() As Byte 'Hier wird der verschlüsselte Text gespeichert

    Using rd As New RijndaelManaged 'Stellt den Rijndael-Algorihmus zur Verfügung

      Using md5 As New MD5CryptoServiceProvider 'Dient zum berechnen eines MD5-Hashes

        Dim key() As Byte = md5.ComputeHash(Encoding.UTF8.GetBytes(Password)) 'Das Passwort wird in ein Byte-Array zerlegt gehasht und dann im Array Key gespeichert

        md5.Clear()   'Leert den MD5CryptoProvider

        rd.Key = key  'Speichert den Key für die Verwendung mit dem Rijndael-Algorithmus

        rd.GenerateIV() 'Hier wird der Initialisierungsvektor berechnet, der für de Verschlüsselung benötigt wird


        Dim ms As New IO.MemoryStream 'Hier wird ein neuer MemoryStream erstellt
        ms.Write(rd.IV, 0, rd.IV.Length) 'und der Initialisierungsvektor hineingeschrieben (die ersten 16 Byte)

        Using cs As New CryptoStream(ms, rd.CreateEncryptor, CryptoStreamMode.Write) 'Nun wird ein neuer Cryptostream aus rd erstellt, welcher in den MemoryStream schreibt

          Dim data() As Byte = System.Text.Encoding.UTF8.GetBytes(Text) 'Der Text wird auch in ein Bytearray zerlegt
          cs.Write(data, 0, data.Length) 'und in den CryptoStream geschrieben
          cs.FlushFinalBlock() 'Garantiert, dass der schreibvorgang abgeschlossen wird

          encdata = ms.ToArray() 'Hier wird der zugrundeliegende MemoryStream in ein ByteArray geschrieben
          cs.Close() 'Schließen und leeren der nicht mehr gebrauchten Instanzen
          rd.Clear()
        End Using
      End Using
    End Using
    Return Convert.ToBase64String(encdata) 'Nun wird der Base64 String der Daten zurück gegeben
  End Function

  Public Function Decrypt(ByVal Text As String, ByVal Password As String) As String
    'Löst eine Exception aus wenn der eingegebene Text bzw das Passwort leer ist oder aus nur Whitespaces besteht
    If IsNullOrWhiteSpace(Text) Then
      If IsNothing(Text) Then
        Throw New ArgumentNullException("Text")
      Else
        Throw New ArgumentException("Argument cannot be empty or whitespace!", "Text")
      End If
      Exit Function
    End If
    If IsNullOrWhiteSpace(Password) Then
      If IsNothing(Password) Then
        Throw New ArgumentNullException("Password")
      Else
        Throw New ArgumentException("Argument cannot be empty or whitespace!", "Password")
      End If
      Exit Function
    End If

    Dim data() As Byte 'Hier wird der entschlüsselte Text gespeichert
    Dim i As Integer
    Using rd As New RijndaelManaged 'Stellt den Rijndael-Algorihmus zur Verfügung
      Dim rijndaelIvLength As Integer = 16 'Die Anzahl der Bytes des Initialisierungsvektor
      Using md5 As New MD5CryptoServiceProvider 'Dient zum berechnen eines MD5-Hashes
        Dim key() As Byte = md5.ComputeHash(Encoding.UTF8.GetBytes(Password)) 'Das Passwort wird in ein Byte-Array zerlegt gehasht und dann im Array Key gespeichert
        md5.Clear() 'Leert den MD5CryptoProvider
        Dim encdata() As Byte = Convert.FromBase64String(Text) 'Wandelt den Base64-String des verschlüsselten Textes in ein Byte Array und speichert dies in encdata
        Using ms As New IO.MemoryStream(encdata) 'Öffnet das ByteArray in einem MemoryStream
          Dim iv(15) As Byte
          ms.Read(iv, 0, rijndaelIvLength) 'Hier wird der Initialisierungsvektor aus dem Memorystream gelesen (die ersten 16 Byte, siehe Verschlüsselung)
          rd.IV = iv
          rd.Key = key 'Hier wird der Key für die weitere Verwendung in der aktuellen Instanz von RijndaelManaged gespeichert
          Dim cs As New CryptoStream(ms, rd.CreateDecryptor, CryptoStreamMode.Read)  'Nun wird ein neuer Cryptostream aus rd erstellt, welcher aus dem MemoryStream liest
          ReDim data(Convert.ToInt32(ms.Length) - Convert.ToInt32(rijndaelIvLength)) 'Hier wird die Größe des Byte Arrays für den entschlüsselten Text festgelegt und dieses neu dimensioniert
          Try
            i = cs.Read(data, 0, data.Length) 'Hier wird der Klartext ausgelesen
            cs.Close() 'Schließen und leeren der nicht mehr gebrauchten Instanzen
          Catch ex As System.Security.Cryptography.CryptographicException
            MsgBox("Wrong Passphrase!")
          Finally
            rd.Clear()
          End Try
        End Using
      End Using
    End Using
    Return System.Text.Encoding.UTF8.GetString(data, 0, i) 'Zum Schluss wird das Byte Array zu einem String umgewandelt
  End Function

  Private Shared Function IsNullOrWhiteSpace(ByVal s As String) As Boolean
    Dim retval As Boolean = True
    If Not String.IsNullOrEmpty(s) Then
      For Each c As Char In s
        If Not Char.IsWhiteSpace(c) Then
          retval = False
        End If
      Next
    End If
    Return retval
  End Function
End Class