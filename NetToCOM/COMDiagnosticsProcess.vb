<ComClass(Process.ClassId, Process.InterfaceId, Process.EventsId)> _
<System.Runtime.InteropServices.ComVisible(True)>
Public Class Process
#Region "COM-GUIDs"
  ' Diese GUIDs stellen die COM-Identität für diese Klasse 
  ' und ihre COM-Schnittstellen bereit. Wenn Sie sie ändern, können vorhandene 
  ' Clients nicht mehr auf die Klasse zugreifen.
  Public Const ClassId As String = "35f8ff17-a69e-4d53-acbf-ea93caf87f98"
  Public Const InterfaceId As String = "af9ff5d0-2b4f-4a19-8d66-32a54dbbbb62"
  Public Const EventsId As String = "7b0c6069-275d-42f5-b7d9-50a70fae01ad"
#End Region
  ' Eine erstellbare COM-Klasse muss eine Public Sub New() 
  ' ohne Parameter aufweisen. Andernfalls wird die Klasse 
  ' nicht in der COM-Registrierung registriert und kann nicht 
  ' über CreateObject erstellt werden.
  Public Sub New()
    MyBase.New()
    _p = New System.Diagnostics.Process
  End Sub

  Public Sub Start()
    _p.Start()
  End Sub

  Public Sub StartWait()
    _p.Start()
    _p.WaitForExit()
  End Sub

  Public Sub Kill()
    _p.Kill()
  End Sub

  Public Sub GetProcessByName(ByVal Name As String)
    _p = Diagnostics.Process.GetProcessesByName(Name)(0)
  End Sub

  Public Sub GetProcessByID(ByVal ID As Integer)
    _p = Diagnostics.Process.GetProcessById(ID)
  End Sub

  Private _p As System.Diagnostics.Process
  Public Property FileName As String
    Get
      Return _p.StartInfo.FileName
    End Get
    Set(value As String)
      _p.StartInfo.FileName = value
    End Set
  End Property

  Public Property WorkingDirectory As String
    Get
      Return _p.StartInfo.WorkingDirectory
    End Get
    Set(value As String)
      _p.StartInfo.WorkingDirectory = value
    End Set
  End Property

  Public ReadOnly Property ExitCode As Integer
    Get
      Return _p.ExitCode
    End Get
  End Property

  Public Property Arguments As String
    Get
      Return _p.StartInfo.Arguments
    End Get
    Set(value As String)
      _p.StartInfo.Arguments = value
    End Set
  End Property

  Public Property CreateNoWindow As Boolean
    Get
      Return _p.StartInfo.CreateNoWindow
    End Get
    Set(value As Boolean)
      _p.StartInfo.CreateNoWindow = value
    End Set
  End Property

  Public Property UseShellExecute As Boolean
    Get
      Return _p.StartInfo.UseShellExecute
    End Get
    Set(value As Boolean)
      _p.StartInfo.UseShellExecute = value
    End Set
  End Property

  Public Property WindowStyle As COMProcessWindowStyle
    Get
      Return CType(_p.StartInfo.WindowStyle, COMProcessWindowStyle)
    End Get
    Set(value As COMProcessWindowStyle)
      _p.StartInfo.WindowStyle = CType(value, ProcessWindowStyle)
    End Set
  End Property

  Public ReadOnly Property HasExited As Boolean
    Get
      Return _p.HasExited
    End Get
  End Property

  Public ReadOnly Property ID As Integer
    Get
      Return _p.Id
    End Get
  End Property

  Public ReadOnly Property Handle As Integer
    Get
      Return _p.Handle.ToInt32
    End Get
  End Property

  <System.Runtime.InteropServices.ComVisible(True)>
  Public Enum COMProcessWindowStyle As Integer
    Normal = 0
    Hidden = 1
    Minimized = 2
    Maximized = 3
  End Enum
End Class