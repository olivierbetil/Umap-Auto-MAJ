VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SeleniumWebDriverUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Rem Did Chrome change their file url and break your code?
' Check for an update: https://stackoverflow.com/a/67996166/5781745

Private ChromeDriver As Selenium.ChromeDriver
Private EdgeDriver As Selenium.EdgeDriver
Private SeleniumFolder As String
Private TempZipFile As String
Private ChromeInit As Boolean, EdgeInit As Boolean

Public Enum dType
    Chrome
    Edge
End Enum

Public Property Get SeleniumFolderPath() As String
    SeleniumFolderPath = SeleniumFolder
End Property

Public Property Let SeleniumFolderPath(ByVal FolderPath As String)
    SeleniumFolder = FolderPath
End Property
    
Public Sub UpdateDriver(ByVal DriverType As dType)

    'URLs to the drivers' home pages to which we can grab the curr versions
    Dim URLPath As String
    Select Case DriverType
    Case dType.Chrome
        URLPath = "https://chromedriver.chromium.org/home"
    Case dType.Edge
        URLPath = "https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/"
    End Select
    
    'Grab the current Version # from the driver's webpage
    Dim Doc As New HTMLDocument, DriverVer As String
    With New MSXML2.XMLHTTP60
        .Open "GET", URLPath
        .Send
        Doc.body.innerHTML = .responseText
    End With
    DriverVer = getCurrentVersion(Doc, DriverType)
    
    DownloadUpdatedDriver DriverVer, DriverType
    ExtractZipAndCopy DriverType

End Sub

' For use in a later project. Not needed at this time
Private Sub InitializeDriver(ByVal DriverType As dType)
    Select Case DriverType
    Case dType.Chrome
        Set ChromeDriver = New Selenium.ChromeDriver
        ChromeDriver.Start
        ChromeInit = True
    Case dType.Edge
        Set EdgeDriver = New Selenium.EdgeDriver
        EdgeDriver.Start
        EdgeInit = True
    End Select
End Sub

Private Function getCurrentVersion(Doc As HTMLDocument, DriverType As dType) As String

    Dim div As HTMLDivElement

    Select Case DriverType
    Case dType.Chrome
        For Each div In Doc.getElementsByTagName("p")
            If div.innerText Like "Latest stable release*" Then
                With New VBScript_RegExp_55.RegExp
                    .Pattern = "ChromeDriver\s([\d\.]+)\b"
                    getCurrentVersion = .Execute(div.innerText)(0).SubMatches(0)
                    Exit Function
                End With
            End If
        Next
    Case dType.Edge
        With New VBScript_RegExp_55.RegExp
            .Pattern = "Version:\s([\d\.]+)"
            For Each div In Doc.getElementsByClassName("module")(0).getElementsByTagName("p")
                If .Test(div.innerText) Then
                    getCurrentVersion = .Execute(div.innerText)(0).SubMatches(0)
                    Exit Function
                End If
            Next
        End With
    End Select

End Function

Private Sub DownloadUpdatedDriver(ByVal CurrVersion As String, DriverType As dType)
    
    Dim URLPath As String
    Select Case DriverType
    Case dType.Chrome
        URLPath = "https://chromedriver.storage.googleapis.com/" & CurrVersion & "/chromedriver_win32.zip"
    Case dType.Edge
        Kill Environ$("LocalAppData") & "\SeleniumBasic\Driver_Notes\*.*"
        URLPath = "https://msedgedriver.azureedge.net/" & CurrVersion & "/edgedriver_win64.zip"
    End Select
    
    Dim FileStream As New ADODB.Stream
    With New MSXML2.XMLHTTP60
        .Open "GET", URLPath
        .Send
        FileStream.Open
        FileStream.Type = adTypeBinary
        FileStream.Write .responseBody
        FileStream.SaveToFile TempZipFile, adSaveCreateOverWrite
        FileStream.Close
    End With
    
End Sub

Private Sub ExtractZipAndCopy(ByVal DriverType As dType)

    Dim FileName As String
    Select Case DriverType
    Case dType.Chrome: FileName = "\chromedriver.exe"
    Case dType.Edge: FileName = "\edgedriver.exe"
    End Select

    'Delete the old WebDriver
    Kill SeleniumFolder & FileName
    
    'Copy the new driver from .zip file to SeleniumBasic folder
    Dim oShell As New Shell
    oShell.Namespace(SeleniumFolder).CopyHere oShell.Namespace(TempZipFile).Items
    
    'Selenium VBA expects 'edgedriver' for edge, but new drivers are named 'msedgedriver'.
    'If we are updating Edge, we need to rename the file
    If DriverType = dType.Edge Then
        Name SeleniumFolder & "msedgedriver.exe" As SeleniumFolder & "edgedriver.exe"
    End If
        
    'Delete the temporary zip file
    Kill TempZipFile

End Sub

Private Sub Class_Initialize()

    ' Set the default file path. Can be modified later using ChromeDriverPath property
    SeleniumFolder = Environ$("LocalAppData") & "\SeleniumBasic\"
    TempZipFile = Environ$("LocalAppData") & "\Temp\WebDriver.zip"

End Sub
