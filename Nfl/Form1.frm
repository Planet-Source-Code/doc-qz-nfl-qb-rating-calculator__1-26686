VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "QB Rating"
   ClientHeight    =   3615
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5745
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   5745
   Begin VB.CommandButton cmdTotal 
      BackColor       =   &H80000009&
      Caption         =   "&Calculate"
      Height          =   375
      Left            =   3600
      TabIndex        =   18
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   3480
      TabIndex        =   6
      Top             =   120
      Width           =   2175
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   1320
         TabIndex        =   16
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1320
         TabIndex        =   10
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1320
         TabIndex        =   9
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1320
         TabIndex        =   8
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1320
         TabIndex        =   7
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Total Yards--------"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Attempts----------"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Completions-------"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Td Passes---------"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "     Picks-------------"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   1215
      End
   End
   Begin VB.TextBox Text5 
      Height          =   615
      Left            =   4320
      TabIndex        =   4
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Quarterback"
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   3255
      Begin VB.ComboBox cboQB 
         Height          =   315
         ItemData        =   "Form1.frx":0ECA
         Left            =   120
         List            =   "Form1.frx":0ECC
         TabIndex        =   3
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Team"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
      Begin VB.ComboBox cboTeams 
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "-This Program is in no way affiliated with the NFL or and Sports Bodies-"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   3360
      Width           =   5535
   End
   Begin VB.Image Image2 
      Height          =   1335
      Left            =   1320
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Rating="
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   2160
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   120
      Top             =   2040
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'code is going in the right direction.
'when a team is selected the QB's are added to cboQB
'So when ever a team is selected the old data in the combo box will be removed
'and the new QB's names displayed. Now an easy way to load the QB's picture when
'he is selected from a combo box would be as following, in the same dir as your
'application place each QB's picture file and give it the exact same name as the
'QB's name that is used in the cboQB combo box. What this will allows us to do is
'easily load each QB picture when the click the cboQB box. Check the cboQB_Click
'Event for more information on how it works.



'I moved the:
    'intcomps = Val(Text2.Text)
    'sngInts = Val(Text1.Text)
    'sngAtt = Val(Text6.Text)
    'sngTtlYds = Val(Text3.Text)
    'sngTds = Val(Text4.Text)
    'sngTotal = Rating(intcomps, sngAtt, sngTds, sngInts, sndTtlYds)
'into a command button named cmdTotal
'Also took variable declarations out of module and put them in the
'General Declarations section of form1
' I would add error checking for the click event(Disable it maybe?)
'until the user has selected a quarterback because if you call the
'function and send it null or blank values, you'll get an error
'I'm not really sure whether a user enters the QB information or if
'the program automaticaly shows it when a QB is selected. So I
'can only help you with what I know. The command button will take
'care of it in both cases, but if the program holds the info, then
'doing it in the cboBQ_Click event would look better.




Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
Dim intcomps As Single
Dim sngInts As Single
Dim sngAtt As Single
Dim sngTtlYds As Single
Dim sngTds As Single
Dim sngTotal As Single

Private Sub cboQB_click()
'When the user clicks a QB we need to load his picture.
'Based on the name in cboQB we will use that to load the proper QB picture file
'This means the picture file needs to be spelled correctly as the QB name
'Example: Jake Plummer is selected so we load Jake Plummer.jpg into the image control
Image2.Picture = LoadPicture(cboQB.Text & ".jpg")
End Sub

Private Sub cboTeams_Click()
    Select Case cboTeams.ListIndex
    Case 0
        'Display team image
        Image1.Picture = LoadPicture("Cardinals.gif")
        'Load qb's into comobo box
        'Clear any old data in combo box
        cboQB.Clear
        cboQB.AddItem "Jake Plummer"
        cboQB.AddItem "Dave Brown"
        cboQB.AddItem "Chris Greisen"
    Case 1
        'Display team image
        Image1.Picture = LoadPicture("falcons.gif")
        'Load qb's into comobo box
        'Clear any old data in combo box
        cboQB.Clear
        cboQB.AddItem "Michael Vick"
        cboQB.AddItem "Chris Chandler"
        cboQB.AddItem "Eric Zeier"
        cboQB.AddItem "Doug Johnson"
    Case 2
        Image1.Picture = LoadPicture("ravens.gif")
        cboQB.Clear
        cboQB.AddItem "Randall Cunningham"
        cboQB.AddItem "Elvis Grbac"
        cboQB.AddItem "Chris Redman"
    Case 3
        Image1.Picture = LoadPicture("bills.gif")
        cboQB.Clear
        cboQB.AddItem "Rob Johnson"
        cboQB.AddItem "Pete Gonzalez"
        cboQB.AddItem "Van Pelt"
    Case 4
        Image1.Picture = LoadPicture("panthers.gif")
        cboQB.Clear
        cboQB.AddItem "Chris Weinke"
        cboQB.AddItem "Dameyune Craig"
        cboQB.AddItem "Jeff lewis"
        cboQB.AddItem "Matt Lytle"
    Case 5
        Image1.Picture = LoadPicture("bears.gif")
        cboQB.Clear
        cboQB.AddItem "Shane Matthews"
        cboQB.AddItem "Jim Miller"
        cboQB.AddItem "Danny Wuerffel"
    Case 6
         Image1.Picture = LoadPicture("bengals.gif")
         cboQB.Clear
         cboQB.AddItem "Scott Covington"
         cboQB.AddItem "Akili Smith"
         cboQB.AddItem "Jon Kitna"
         cboQB.AddItem "Scott Mitchell"
    Case 7
         Image1.Picture = LoadPicture("browns.gif")
         cboQB.Clear
         cboQB.AddItem "Tim Couch"
         cboQB.AddItem "Kelly Holcomb"
         cboQB.AddItem "Kevin Thompson"
    Case 8
         Image1.Picture = LoadPicture("cowboys.gif")
         cboQB.Clear
         cboQB.AddItem "Quincy Carter"
         cboQB.AddItem "clint stoerner"
         cboQB.AddItem "Anthony Wright"
    Case 9
         Image1.Picture = LoadPicture("broncos.gif")
         cboQB.Clear
         cboQB.AddItem "Steve Beurlein"
         cboQB.AddItem "Gus Frerotte"
         cboQB.AddItem "Brian Griese"
    Case 10
         Image1.Picture = LoadPicture("lions.gif")
         
         cboQB.Clear
         cboQB.AddItem "Charlie Batch"
         cboQB.AddItem "Jim Harbaugh"
         cboQB.AddItem "Cory Sauter"
    Case 11
         Image1.Picture = LoadPicture("packers.gif")
         cboQB.Clear
         cboQB.AddItem " Brett Favre"
         cboQB.AddItem "Doug Pederson"
         cboQB.AddItem "Billy Joe Tolliver"
    Case 12
         Image1.Picture = LoadPicture("colts.gif")
         cboQB.Clear
         cboQB.AddItem "Peyton Manning"
         cboQB.AddItem "Mark Rypien"
         cboQB.AddItem "Billy Joe Hobert"
     Case 13
         Image1.Picture = LoadPicture("jaguars.gif")
         cboQB.Clear
         cboQB.AddItem "Mark Brunell"
         cboQB.AddItem "Jamie Martin"
         cboQB.AddItem "Jonathan Quinn"
    Case 14
         Image1.Picture = LoadPicture("chiefs.gif")
         cboQB.Clear
         cboQB.AddItem "Bubby Brister"
         cboQB.AddItem "Trent Green"
         cboQB.AddItem "Todd Collins"
    Case 15
          Image1.Picture = LoadPicture("dolphins.gif")
          cboQB.Clear
          cboQB.AddItem "Jay Fiedler"
          cboQB.AddItem "Ray Lucas"
          cboQB.AddItem "Josh Heupel"
    Case 16
          Image1.Picture = LoadPicture("vikings.gif")
          cboQB.Clear
          cboQB.AddItem "Daunte Culpepper"
          cboQB.AddItem "Todd Bouman"
    Case 17
         Image1.Picture = LoadPicture("patriots.gif")
         cboQB.Clear
         cboQB.AddItem "Drew Bledsoe"
         cboQB.AddItem "Tom Brady"
         cboQB.AddItem "Michael Bishop"
         cboQB.AddItem "Damon Huard"
    Case 18
         Image1.Picture = LoadPicture("saints.gif")
         cboQB.Clear
         cboQB.AddItem "Jeff Blake"
         cboQB.AddItem "Aaron Brooks"
         cboQB.AddItem "Jake Delhomme"
    Case 19
          Image1.Picture = LoadPicture("giants.gif")
          cboQB.Clear
          cboQB.AddItem "Kerry Collins"
          cboQB.AddItem "Jason Garrett"
          cboQB.AddItem "Jesse Palmer"
    Case 20
          Image1.Picture = LoadPicture("jets.gif")
          cboQB.Clear
          cboQB.AddItem "Chad Pennington"
          cboQB.AddItem "Vinny Testaverde"
          
    Case 21
          Image1.Picture = LoadPicture("raiders.gif")
          cboQB.Clear
          cboQB.AddItem "Rich Gannon"
          cboQB.AddItem "Marques Tuiasosopo"
          cboQB.AddItem "Bobby Hoying"
    Case 22
          Image1.Picture = LoadPicture("eagles.gif")
          cboQB.Clear
          cboQB.AddItem "Koy Detmer"
          cboQB.AddItem "Donovan McNabb"
          cboQB.AddItem "Ron Powlus"
    Case 23
          Image1.Picture = LoadPicture("steelers.gif")
          cboQB.Clear
          cboQB.AddItem "Kordell Stewart"
          cboQB.AddItem "Kent Graham"
          cboQB.AddItem "Tee Martin"
    Case 24
          Image1.Picture = LoadPicture("chargers.gif")
          cboQB.Clear
          cboQB.AddItem "Doug Flutie"
          cboQB.AddItem "Drew Brees"
          
    Case 25
          Image1.Picture = LoadPicture("49ers.gif")
          cboQB.Clear
          cboQB.AddItem "Jeff Garcia"
          cboQB.AddItem "Rick Mirer"
          cboQB.AddItem "Tim Rattay"
    Case 26
           Image1.Picture = LoadPicture("seahawks.gif")
           cboQB.Clear
           cboQB.AddItem "Trent Dilfer"
           cboQB.AddItem "Hatt Hasslebeck"
           cboQB.AddItem "Josh Booty"
           cboQB.AddItem "Brock Huard"
    Case 27
          Image1.Picture = LoadPicture("rams.gif")
          cboQB.Clear
          cboQB.AddItem "Kurt Warner"
          cboQB.AddItem "Marc Bulger"
          cboQB.AddItem "Joe Germaine"
          cboQB.AddItem "Paul Justin"
    Case 28
           Image1.Picture = LoadPicture("bucs.gif")
           cboQB.Clear
           cboQB.AddItem "Brad Johnson"
           cboQB.AddItem "Shaun King"
           cboQB.AddItem "Ryan Leaf"
           cboQB.AddItem "Joe Hamilton"
    Case 29
          Image1.Picture = LoadPicture("titans.gif")
          cboQB.Clear
          cboQB.AddItem "Steve McNair"
          cboQB.AddItem "Neil O'Donnell"
          cboQB.AddItem "Billy Voleck"
    Case 30
          Image1.Picture = LoadPicture("redskins.gif")
          cboQB.Clear
          cboQB.AddItem "Jeff George"
          cboQB.AddItem "Cade McNown"
          cboQB.AddItem "Todd Husak"
    End Select
    Image1.Visible = True

End Sub

Private Sub cmdTotal_Click()
   intcomps = Val(Text2.Text)
    sngInts = Val(Text1.Text)
     sngAtt = Val(Text6.Text)
   sngTtlYds = Val(Text3.Text)
      sngTds = Val(Text4.Text)
     sngTotal = Rating(intcomps, sngAtt, sngTds, sngInts, sngTtlYds)
   Text5.Text = sngTotal
End Sub

Private Sub Form_Load()
'This will load all needed data into combo box
Call PopulateComboBox
End Sub

Public Sub PopulateComboBox()
'This sub is used to add all needed data to our combo box
cboTeams.AddItem "Arizona Cardinals"
cboTeams.AddItem "Atlanta Falcons"
cboTeams.AddItem "Baltimore Ravens"
cboTeams.AddItem "Buffalo Bills"
cboTeams.AddItem "Carolina Panthers"
cboTeams.AddItem "Chicago Bears"
cboTeams.AddItem "Cincinnati Bengals"
cboTeams.AddItem "Cleveland Browns"
cboTeams.AddItem "Dallas Cowboys"
cboTeams.AddItem "Denver Broncos"
cboTeams.AddItem "Detroit Lions"
cboTeams.AddItem "Green Bay Packers"
cboTeams.AddItem "Indianapolis Colts"
cboTeams.AddItem "Jacksonville Jaguars"
cboTeams.AddItem "Kansas City Chiefs"
cboTeams.AddItem "Miami Dolphins"
cboTeams.AddItem "Minnesota Vikings"
cboTeams.AddItem "New England Patriots"
cboTeams.AddItem "New Orleans Saints"
cboTeams.AddItem "NY Giants"
cboTeams.AddItem "NY Jets"
cboTeams.AddItem "Oakland Raiders"
cboTeams.AddItem "Philadelphia Eagles"
cboTeams.AddItem "Pittsburgh Steelers"
cboTeams.AddItem "San Diego Chargers"
cboTeams.AddItem "San Francisco 49'ers"
cboTeams.AddItem "Seattle Seahawks"
cboTeams.AddItem "St. Louis Rams"
cboTeams.AddItem "Tampa Bay Buccaneers"
cboTeams.AddItem "Tennessee Titans"
cboTeams.AddItem "Washington Redskins"
End Sub


 



