Attribute VB_Name = "Module1"
'Following two functions are used to move the windows around by clicking directly on
'the window in abscence of the standard window title bars
Public Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ReleaseCapture Lib "User32" () As Long

'creates a region to make form non rectangular
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long

'assigns a region to the form
Public Declare Function SetWindowRgn Lib "User32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

'To delete the regions to free up resources when exiting applicaton
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

'An alternative to the timer (really slow on win98)
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Another choice for region Shape
Public Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

'Used for "always on top" function to work
Public Declare Sub SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

'Standard Constants Used In Above Functions
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_FRAMECHANGED = &H20

