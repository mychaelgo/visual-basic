VERSION 5.00
Begin VB.Form RoOm_StATus 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "RoOm_StATus"
   ClientHeight    =   4695
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   7905
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   7905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   4335
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   1575
      Begin VB.ListBox List1 
         Height          =   3765
         ItemData        =   "RoOm_StATus.frx":0000
         Left            =   240
         List            =   "RoOm_StATus.frx":003D
         TabIndex        =   21
         Top             =   480
         Width           =   1095
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   "D:\ashik\HMS\hotel2.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "room"
         Top             =   4680
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Label Label1 
         Caption         =   "Room Number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4335
      Left            =   1920
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.CommandButton Command1 
         Caption         =   "Close"
         Height          =   375
         Left            =   3000
         TabIndex        =   2
         Top             =   4440
         Width           =   1095
      End
      Begin VB.Data Data2 
         Caption         =   "Data2"
         Connect         =   "Access"
         DatabaseName    =   "D:\ashik\HMS\hotel2.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   1000
         Left            =   500
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "checkin"
         Top             =   4680
         Width           =   1140
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Close"
         Height          =   375
         Left            =   3000
         TabIndex        =   1
         Top             =   3840
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Room Number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   19
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Age"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Sex"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Phone"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "Arrival Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Duration"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   3720
         Width           =   1095
      End
      Begin VB.Label Lroom 
         Height          =   255
         Left            =   2040
         TabIndex        =   11
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Lnm 
         Height          =   255
         Left            =   2040
         TabIndex        =   10
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Lage 
         Height          =   255
         Left            =   2040
         TabIndex        =   9
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Lsex 
         Height          =   255
         Left            =   2040
         TabIndex        =   8
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Laddr 
         Height          =   255
         Left            =   2040
         TabIndex        =   7
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label Label15 
         Height          =   255
         Left            =   2160
         TabIndex        =   6
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Label Ldate 
         Height          =   255
         Left            =   2040
         TabIndex        =   5
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Label Ldurasi 
         Height          =   255
         Left            =   2040
         TabIndex        =   4
         Top             =   3720
         Width           =   735
      End
      Begin VB.Label Lphone 
         Height          =   255
         Left            =   2040
         TabIndex        =   3
         Top             =   2760
         Width           =   735
      End
   End
End
Attribute VB_Name = "RoOm_StATus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
