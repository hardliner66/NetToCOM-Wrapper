Option Strict On
'<ComClass(MemoryStream.ClassId, MemoryStream.InterfaceId, MemoryStream.EventsId)> _
Public Class MemoryStream
#Region "COM-GUIDs"
  ' Diese GUIDs stellen die COM-Identität für diese Klasse 
  ' und ihre COM-Schnittstellen bereit. Wenn Sie sie ändern, können vorhandene 
  ' Clients nicht mehr auf die Klasse zugreifen.
  Public Const ClassId As String = "4190dbcd-a36b-4ee7-b761-e25a7af43a6b"
  Public Const InterfaceId As String = "1315bf4d-ca3c-429f-813b-55f69e9c8761"
  Public Const EventsId As String = "b52ec8cd-c230-4242-85d6-7e85a76c6a80"
#End Region
  ' Eine erstellbare COM-Klasse muss eine Public Sub New() 
  ' ohne Parameter aufweisen. Andernfalls wird die Klasse 
  ' nicht in der COM-Registrierung registriert und kann nicht 
  ' über CreateObject erstellt werden.
  Public Sub New()
    MyBase.New()
    ms = New IO.MemoryStream(Nothing, True)
  End Sub

  Private ms As IO.MemoryStream

  Public Function GetBuffer() As Byte()
    Return ms.GetBuffer
  End Function

  Public Sub CreateStream(ByRef buffer() As Byte)
    ms.Dispose()
    ms = New IO.MemoryStream(buffer, True)
  End Sub

  Public Sub Write(ByRef buffer() As Byte, ByVal offset As Integer, ByVal count As Integer)
    ms.Write(buffer, offset, count)
  End Sub

  Public Function Read(ByRef buffer() As Byte, ByVal offset As Integer, ByVal count As Integer) As Integer
    Return ms.Read(buffer, offset, count)
  End Function

  Public Sub Dispose()
    ms.Dispose()
  End Sub

  Public Function ToMemoryStream() As IO.MemoryStream
    Return ms
  End Function
End Class
