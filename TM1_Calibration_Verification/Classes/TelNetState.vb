Public Class TelNetState
    Private mPortState As Boolean
    Private mUserLogin As Boolean
    Private mCloseRead As Boolean
    Private mConnected As Boolean
    Private mConnecting As Boolean
    Private mFailed As Boolean
    Private mTimedOut As Boolean

    Public Property PortState() As Boolean        
        Get
            Return mPortState
        End Get
        Set(ByVal Value As Boolean)
            mPortState = Value
        End Set
    End Property
    Public Property UserLogin() As Boolean
        Get
            Return mUserLogin
        End Get
        Set(ByVal Value As Boolean)
            mUserLogin = Value
        End Set
    End Property
    Public Property CloseRead() As Boolean
        Get
            Return mCloseRead
        End Get
        Set(ByVal Value As Boolean)
            mCloseRead = Value
        End Set
    End Property
    Public Property Connected() As Boolean
        Get
            Return mConnected
        End Get
        Set(ByVal Value As Boolean)
            mConnected = Value
        End Set
    End Property
    Public Property Connecting() As Boolean
        Get
            Return mConnecting
        End Get
        Set(ByVal Value As Boolean)
            mConnecting = Value
        End Set
    End Property
    Public Property Failed() As Boolean
        Get
            Return mFailed
        End Get
        Set(ByVal Value As Boolean)
            mFailed = Value
        End Set
    End Property
    Public Property TimedOut() As Boolean
        Get
            Return mTimedOut
        End Get
        Set(ByVal Value As Boolean)
            mTimedOut = Value
        End Set
    End Property
End Class
