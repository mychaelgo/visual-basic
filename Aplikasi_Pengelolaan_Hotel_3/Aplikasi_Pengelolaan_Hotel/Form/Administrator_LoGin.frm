VERSION 5.00
Begin VB.Form ADMINSTRATOR_LOGIN 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "ADMINSTRATOR_LOGIN"
   ClientHeight    =   5085
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   7695
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "New user registration"
      Height          =   2535
      Left            =   1200
      TabIndex        =   21
      Top             =   480
      Width           =   4455
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   960
         TabIndex        =   29
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   960
         TabIndex        =   28
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   3120
         TabIndex        =   27
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   960
         TabIndex        =   26
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox Text9 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   25
         Top             =   1800
         Width           =   975
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Add"
         Height          =   375
         Left            =   3120
         TabIndex        =   24
         Top             =   960
         Width           =   855
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Clear"
         Height          =   375
         Left            =   3120
         TabIndex        =   23
         Top             =   1440
         Width           =   855
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Close"
         Height          =   375
         Left            =   3120
         TabIndex        =   22
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Name"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label9 
         Caption         =   "Address"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "Phone"
         Height          =   255
         Left            =   2400
         TabIndex        =   32
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "UserId"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "Password"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   1800
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   720
      TabIndex        =   8
      Top             =   480
      Width           =   4455
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1320
         TabIndex        =   16
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Add User"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   2040
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Edit User"
         Height          =   375
         Left            =   1200
         TabIndex        =   14
         Top             =   2040
         Width           =   975
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Delete User"
         Height          =   375
         Left            =   2280
         TabIndex        =   13
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1320
         TabIndex        =   12
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1320
         PasswordChar    =   "*"
         TabIndex        =   11
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Clear"
         Height          =   375
         Left            =   3360
         TabIndex        =   10
         Top             =   2040
         Width           =   975
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Update"
         Height          =   375
         Left            =   2280
         TabIndex        =   9
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "User List"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "User Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1320
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1695
      Left            =   720
      TabIndex        =   0
      Top             =   3240
      Width           =   4455
      Begin VB.TextBox Text3 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1560
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox Text4 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1560
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Edit"
         Height          =   375
         Left            =   3240
         TabIndex        =   2
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Update"
         Height          =   375
         Left            =   3240
         TabIndex        =   1
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Administrator ID Management"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   2655
      End
      Begin VB.Label Label6 
         Caption         =   "Admin UserId"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Admin Password"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   1335
      End
   End
   Begin VB.Label Label1 
      Caption         =   "ADMINISTRATOR AREA"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   1440
      TabIndex        =   20
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "ADMINSTRATOR_LOGIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
