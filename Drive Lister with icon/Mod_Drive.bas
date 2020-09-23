Attribute VB_Name = "Mod_Drive"
'
'
'
'
'
'
'
'
'
'
'
'
'
'               Author : Rajeev P
'               Date   : 9/9/05
'               Email  : Rajeev_punnalil@hotmail.com
'
'               Hi friends,this is my second submission to planet
'source code . I developed this code as part for my media player to detect cd drives
'but i thought it would be helpful if i could list all drives. so created this.
'you should'nt have any difficulty understanding this 'cause it's straight forward.so
'i havent commented it much ok ! have a nice time and enjoy !
'
'           If u have any suggestions or queries,please feel free to mail me
'
'
'
'
'
'
'
'
'
'
'
'

Option Explicit

' Api Declarations
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

'Drive type constants
Private Const DRIVE_REMOVABLE = 2
Private Const DRIVE_FIXED = 3
Private Const DRIVE_REMOTE = 4
Private Const DRIVE_CDROM = 5
Private Const DRIVE_RAMDISK = 6


Public Enum DriveTypeConst
    [REMOVABLE] = 2
    [Fixed] = 3
    [REMOTE] = 4
    [CDROM] = 5
    [RAMDISK] = 6
End Enum
Public Function GetDrives(ByRef LstBox As ListBox, Optional DriveModel As DriveTypeConst) As Integer
    Dim ret As Long, AllDrives As String, IsolatedDrive As String, Posn As Long, DriveType As Long
    Dim NumDrives As Integer
    LstBox.Clear
    AllDrives = Space(64)
    ret = GetLogicalDriveStrings(Len(AllDrives), AllDrives)
    AllDrives = Left(AllDrives, ret)
    Do
    Posn = InStr(AllDrives, Chr(0))
    If Posn Then
            IsolatedDrive = Left(AllDrives, Posn)
            AllDrives = Mid(AllDrives, Posn + 1, Len(AllDrives))
            DriveType& = GetDriveType(IsolatedDrive)
            If DriveModel = 0 Then
                  LstBox.AddItem Mid(IsolatedDrive, 1, 2) & "      " & GetDetails(DriveType)
                  NumDrives = NumDrives + 1
            Else
                If DriveType = DriveModel Then
                    LstBox.AddItem Mid(IsolatedDrive, 1, 2) & "      " & GetDetails(DriveType)
                    NumDrives = NumDrives + 1
                End If
            End If
            
          End If
      Loop Until AllDrives = ""
      GetDrives = NumDrives
End Function

Private Function GetDetails(DriveType As Long) As String
    Select Case DriveType
        Case DRIVE_CDROM:
                GetDetails = "CD DRIVE"
        Case DRIVE_FIXED:
                GetDetails = "FIXED DISK DRIVE"
        Case DRIVE_RAMDISK:
                GetDetails = "RAMDISK DRIVE"
        Case DRIVE_REMOTE:
                GetDetails = "REMOTE DRIVE"
        Case DRIVE_REMOVABLE:
                GetDetails = "REMOVABLE DISK DRIVE"
        Case Else:
                GetDetails = "UNRECOGNISED FILE TYPE"
                
    End Select
End Function




