VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "PathFinding Simple"
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   206
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   585
   End
   Begin VB.Label lb 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      Height          =   195
      Index           =   0
      Left            =   315
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Menu mOption 
      Caption         =   "Option"
      Begin VB.Menu mShowLabels 
         Caption         =   "Show Labels"
      End
      Begin VB.Menu mMsgBox 
         Caption         =   "MsgBox after each move"
      End
      Begin VB.Menu mSet 
         Caption         =   "Set Map"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mHelp 
      Caption         =   "Help"
      Begin VB.Menu mControls 
         Caption         =   "Controls"
      End
      Begin VB.Menu mAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' made by Mephisto... the shortest pathfinding algorithm ive seen

'pos is a two dimensional array. It is the core of all variables. It holds the map values
'1 is start, 2 is end, 9 is wall
'Then, all values above 100 are extensions. 100 is the first extension from the Start. 101 is
'the second extension. and so on and so forth. You will see what i mean in the code
Dim pos() As Integer
'cell size in pixels
Dim cz As Integer

Dim Up As Integer, Across As Integer
'used for loops
Dim i As Integer, j As Integer
'Start X and Start Y
Dim SX As Integer, SY As Integer
'Finish X and Finish Y
Dim FX As Integer, FY As Integer
'Its an important variable. Dont know why i called it tim :D should be like times or something
'but anyway.. it always points to the squares taht are to be extended next turn
Dim tim As Integer
'Boolean that stores whether the tager is reachable or not
Dim reachable As Boolean

'These help to determine if target is reachable or not.
'Its a little programming trick that makes my life easier when to determine when the
'squares cannot be extended more. I add up values of all squares on the map...
'Then i store it to lastSum. The next sum i get, i compare to the LastSum. If they are equal,
'nothing has changed. So no extensions took place, its finished. If there is difference then
'something got extended etc, and its still working
Dim Sum As Long
Dim LastSum As Long

'used in the end to determine the shortest path back to start
Dim LastX As Integer
Dim LastY As Integer

'a boolean to use, in the end to determine shortest path back to start
Dim did As Boolean

'if user wants to have msgbox each turn or not
Dim ShowMsgBox As Boolean, ShowLabels As Boolean
'to be sure we arent using some undeclared variables, type mismatches etc
Option Explicit

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'Pretty straight forward... Key is pressed, then find out which key and do accordingly
If KeyCode = vbKeyF Then
Call FindPath

ElseIf KeyCode = vbKeyR Then
ReInitialize

ElseIf KeyCode = vbKeyG Then
ReInitialize
GenerateLandScape

End If
End Sub

Private Sub Form_Load()
Randomize
Form1.BackColor = vbWhite

ReDim pos(75, 50)
cz = 15
'Default variables
Across = UBound(pos, 1)
Up = UBound(pos, 2)

SX = -1
SY = -1
FX = -1
FY = -1

'Move the label up so it will be just below the squares
Label1.Move 1, Up * cz + 20
'Draw the grid... lines
DrawGrid
End Sub

Sub DrawGrid()
'Pretty straight forward... we go across and vertically drawing lines
For i = 1 To Across
Form1.Line (i * cz, 0)-(i * cz, Up * cz)
Next i

For i = 1 To Up
Form1.Line (0, i * cz)-(Across * cz, i * cz)
Next i
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'We always let the user know where he is... just so its userfriendly. I also used with debugging
Label1.Caption = "X: " & Int(X / cz) & " Y: " & Int(Y / cz)
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Crucial part... its the User SETUP part... When he clicks the button

'We test > Was the click INSIDE the squares ?
If X < Across * cz And Y < Up * cz Then
        If Button = 1 And Shift = 0 Then
        'START POSITION
            'This If is here just in case the user accidentally puts Start position twice
            'If the StartX and StartY are still -1, that means this is the first start he made yet,
            'we just make the start, otherwise, we delete the last start, and make a new one
            If SX = -1 And SY = -1 Then
            pos(Int(X / cz), Int(Y / cz)) = 1
            SX = Int(X / cz)
            SY = Int(Y / cz)
            Form1.Line (Int(X / cz) * cz, Int(Y / cz) * cz)-(Int(X / cz) * cz + cz, Int(Y / cz) * cz + cz), vbGreen, BF
            Else
            pos(SX, SY) = 0
            Form1.Line (SX * cz, SY * cz)-(SX * cz + cz, SY * cz + cz), vbWhite, BF
            SX = Int(X / cz)
            SY = Int(Y / cz)
            pos(Int(X / cz), Int(Y / cz)) = 1
            Form1.Line (Int(X / cz) * cz, Int(Y / cz) * cz)-(Int(X / cz) * cz + cz, Int(Y / cz) * cz + cz), vbGreen, BF
            End If
        ElseIf Button = 2 And Shift = 0 Then
        'END POSITION
            'same thing as with Start postion
            If FX = -1 And FY = -1 Then
            pos(Int(X / cz), Int(Y / cz)) = 2
            FX = Int(X / cz)
            FY = Int(Y / cz)
            Form1.Line (Int(X / cz) * cz, Int(Y / cz) * cz)-(Int(X / cz) * cz + cz, Int(Y / cz) * cz + cz), vbRed, BF
            Else
            pos(FX, FY) = 0
            Form1.Line (FX * cz, FY * cz)-(FX * cz + cz, FY * cz + cz), vbWhite, BF
            FX = Int(X / cz)
            FY = Int(Y / cz)
            pos(Int(X / cz), Int(Y / cz)) = 2
            Form1.Line (FX * cz, FY * cz)-(FX * cz + cz, FY * cz + cz), vbRed, BF
            End If
        
        ElseIf Button = 1 And Shift = 1 Then
        'WALL
        pos(Int(X / cz), Int(Y / cz)) = 9
        Form1.Line (Int(X / cz) * cz, Int(Y / cz) * cz)-(Int(X / cz) * cz + cz, Int(Y / cz) * cz + cz), vbBlack, BF
        ElseIf Button = 2 And Shift = 1 Then
        'Delete wall
        pos(Int(X / cz), Int(Y / cz)) = 0
        Form1.Line (Int(X / cz) * cz, Int(Y / cz) * cz)-(Int(X / cz) * cz + cz, Int(Y / cz) * cz + cz), vbWhite, BF
        End If
'and we restart grid, because the squares are slightly bigger and the grid then dissapeares...
DrawGrid
End If

End Sub

Sub FindPath()
'Initialization phase
tim = 0
pos(SX, SY) = 100 + tim

'reset reachable
reachable = False

'Do while reachable is false... if its set true in progress, we jump out
'If the path is decided unreachable in process, we will use exit sub. Not proper,
'but faster ;-)
Do While reachable = False

'we loop through all squares
For j = 0 To Up - 1
    For i = 0 To Across - 1
    DoEvents
        'If they are to be extended, the pointer TIM is on them
        If pos(i, j) = 100 + tim Then
        'The part is to be extended, so do it
            'We have to make sure that there is a pos(i+1,j) BEFORE we actually use it,
            'because then we get error... If the square is on side, we dont test for this one!
            If i < Across - 1 Then
                'If there isnt a wall, or any other... thing
                If pos(i + 1, j) = 0 Then
                'Expand it, and make its pos equal to tim+1, so the next time we make this loop,
                'It will exapand that square too! This is crucial part of the program
                pos(i + 1, j) = 100 + tim + 1
                'make the square
                Form1.Line ((i + 1) * cz, j * cz)-((i + 1) * cz + cz, j * cz + cz), vbCyan, BF
                'This is just to give the square a number.., its just so it looks nice, this is
                'completely unnecessary i just htought it looks cool... :D
                'we load a new label
                If ShowLabels = True Then
                    Load lb(lb.UBound + 1)
                    'make it visible
                    lb(lb.UBound).Visible = True
                    'we move it
                    lb(lb.UBound).Left = (i + 1) * cz
                    lb(lb.UBound).Top = j * cz
                    'and we set its caption
                    lb(lb.UBound).Caption = tim + 1
                End If
                
                ElseIf pos(i + 1, j) = 2 Then
                'If the position is no 0 but its 2 (FINISH) then Reachable = true!!! We found end
                reachable = True
                End If
            End If
        
            'This is the same as the last one, as i said a lot of copy paste work and editing that
            'This is simply another side that we have to test for... so instead of i+1 we have i-1
            'Its actually pretty same then... I wont comment it therefore, because its only repeating
            'same thing with minor changes to check sides
            If i > 0 Then
                If pos((i - 1), j) = 0 Then
                pos(i - 1, j) = 100 + tim + 1
                Form1.Line ((i - 1) * cz, j * cz)-((i - 1) * cz + cz, j * cz + cz), vbCyan, BF
                
                If ShowLabels = True Then
                    Load lb(lb.UBound + 1)
                    lb(lb.UBound).Visible = True
                    lb(lb.UBound).Left = (i - 1) * cz
                    lb(lb.UBound).Top = j * cz
                    lb(lb.UBound).Caption = tim + 1
                End If
                
                ElseIf pos(i - 1, j) = 2 Then
                'MsgBox "Reachable"
                reachable = True
                End If
            End If
        
            If j < Up - 1 Then
                If pos(i, j + 1) = 0 Then
                pos(i, j + 1) = 100 + tim + 1
                Form1.Line (i * cz, (j + 1) * cz)-(i * cz + cz, (j + 1) * cz + cz), vbCyan, BF
                
                If ShowLabels = True Then
                    Load lb(lb.UBound + 1)
                    lb(lb.UBound).Visible = True
                    lb(lb.UBound).Left = i * cz
                    lb(lb.UBound).Top = (j + 1) * cz
                    lb(lb.UBound).Caption = tim + 1
                End If
                
                ElseIf pos(i, j + 1) = 2 Then
                'MsgBox "Reachable"
                reachable = True
                End If
            End If
        
            If j > 0 Then
                If pos(i, j - 1) = 0 Then
                    pos(i, j - 1) = 100 + tim + 1
                    Form1.Line (i * cz, (j - 1) * cz)-(i * cz + cz, (j - 1) * cz + cz), vbCyan, BF
                    Load lb(lb.UBound + 1)
                    
                    If ShowLabels = True Then
                    lb(lb.UBound).Visible = True
                    lb(lb.UBound).Left = i * cz
                    lb(lb.UBound).Top = (j - 1) * cz
                    lb(lb.UBound).Caption = tim + 1
                    End If
                    
                    ElseIf pos(i, j - 1) = 2 Then
                    'MsgBox "Reachable"
                    reachable = True
                End If
            End If
    End If
    Next i
Next j

'If the reachable is STILL false, then
If reachable = False Then
    'reset sum
    Sum = 0
    For j = 0 To Up - 1
        For i = 0 To Across - 1
        'we add up ALL the squares
        Sum = Sum + pos(i, j)
        Next i
    Next j
    
    'Now if the sum is euqal to the last sum, its not reachable, if it isnt, then we store
    'sum to lastsum
    If Sum = LastSum Then
    DrawGrid
    MsgBox "Destination not reachable"
    Exit Sub
    Else
    LastSum = Sum
    End If
End If

'we increase the pointer to point to the next squares to be expanded
tim = tim + 1

'If the user wants to have msgbox displayed each turn, we make it happen along with a simple
'Psuedo code description so he can see the psuedo code at work and examine it carefully
If ShowMsgBox = True Then
MsgBox "Every Square that is on outside, expend it to each direction if possible. Then change the counter tim to point to all next squares. They are to be expanded next"
End If

Loop

'we again redraw the grid
DrawGrid
'Its reachable, because we are out of the loop
MsgBox "Reachable on " & tim & " moves"

'We work backwards to find the way...
LastX = FX
LastY = FY

'The following code may be a little bit confusing but ill try my best to explain it.
'We are working backwards to find ONE of the shortest ways back to Start.
'So we repeat the loop until the LastX and LastY arent in start. Look in the code to see
'how LastX and LasY change
Do While LastX <> SX Or LastY <> SY
    DoEvents
    'We decrease tim by one, and then we are finding any adjacent square to the final one, that
    'has that value. So lets say the tim would be 5, because it takes 5 steps to get to the target.
    'Now everytime we decrease that, so we make it 4, and we look for any adjacent square that has
    'that value. When we find it, we just color it yellow as for the solution
    tim = tim - 1
    'reset did to false
    did = False
    
    'If we arent on edge
    If LastX < Across Then
        'check the square on the right of the solution. Is it a tim-1 one? or just a blank one
        If pos(LastX + 1, LastY) = 100 + tim Then
            'if it, then make it yellow, and change did to true
            LastX = LastX + 1
            Form1.Line (LastX * cz, LastY * cz)-(LastX * cz + cz, LastY * cz + cz), vbYellow, BF
            did = True
        End If
    End If
    
    'This will then only work if the previous part didnt execute, and did is still false. THen
    'we want to check another square, the on left. Is it a tim-1 one ?
    If did = False Then
        If LastX > 0 Then
            If pos(LastX - 1, LastY) = 100 + tim Then
                LastX = LastX - 1
                Form1.Line (LastX * cz, LastY * cz)-(LastX * cz + cz, LastY * cz + cz), vbYellow, BF
                did = True
            End If
        End If
    End If
    
    'We check the one below it
    If did = False Then
        If LastY < Up Then
            If pos(LastX, LastY + 1) = 100 + tim Then
                LastY = LastY + 1
                Form1.Line (LastX * cz, LastY * cz)-(LastX * cz + cz, LastY * cz + cz), vbYellow, BF
                did = True
            End If
        End If
    End If
    
    'And above it. One of these have to be it, since we have found the solution, we know that already
    'there is a way back.
    If did = False Then
        If LastY > 0 Then
            If pos(LastX, LastY - 1) = 100 + tim Then
                LastY = LastY - 1
                Form1.Line (LastX * cz, LastY * cz)-(LastX * cz + cz, LastY * cz + cz), vbYellow, BF
            End If
        End If
    End If
    
    'Now we loop back and decrease tim, and look for the next square with lower value
Loop

'Because this method also colors yellow the Starting square, we just color it back to green
Form1.Line (LastX * cz, LastY * cz)-(LastX * cz + cz, LastY * cz + cz), vbGreen, BF
'and we draw grid. we are done
DrawGrid
End Sub

Sub ReInitialize()
'self explanatory
For i = 1 To lb.UBound
Unload lb(i)
Next i

For i = 0 To Across
    For j = 0 To Up
    pos(i, j) = 0
    Next j
Next i

SX = -1
SY = -1
FX = -1
FY = -1

Form1.Cls
DrawGrid
End Sub

Sub GenerateLandScape()
For i = 0 To Across - 1
    For j = 0 To Up - 1
    'a very cheap and effective method for this purpose, loop through all squares. Then
    'generate a number 1-3, if it is a 1, then make it a wall... I didnt go any further
    'then that, but of course it is doable
    If Int((Rnd * 3) + 1) = 1 Then
        pos(i, j) = 9
        Form1.Line (i * cz, j * cz)-(i * cz + cz, j * cz + cz), vbBlack, BF
    End If
    Next j
Next i
End Sub

Private Sub mAbout_Click()
'Do i really need to comment this part too ? :p
MsgBox "A simple presentation of how to make a pathfinding program. This is about the simplest" & vbCrLf & _
        "pathfinding program ive ever seen. I wanted to integrate pathfinding to my games, but all codes that i got" & vbCrLf & _
        "were really big, 5 moduls and a lot of API's so i decided to take a day or two writing my own algorithm." & vbCrLf & _
        "and this is the outcome. The code may seem a little bit too big, but it is mainly because i tried to make it userfriendly, " & vbCrLf & _
        "there is a lot of copy and paste code with 5 changes, because we are always checking 4 sides, and we need code for each side separatly." & vbCrLf & _
        "anyway, enjoy :D i tried to make it as simple as i could." & vbCrLf & "Made by Mephisto"

End Sub

Private Sub mControls_Click()
MsgBox "Left Click : Start Position" & vbCrLf & "Right Click : Finish Position" & vbCrLf & _
       "Shift+Left Click : Wall" & vbCrLf & "Shift+Right Click : Erase Wall" & vbCrLf & _
       "F : Find Path" & vbCrLf & "R : Erase all" & vbCrLf & "G : Erase all and generate new terrain"
End Sub

Private Sub mMsgBox_Click()
If mMsgBox.Checked = False Then
ShowMsgBox = True
mMsgBox.Checked = True
Else
ShowMsgBox = False
mMsgBox.Checked = False
End If
End Sub

Private Sub mSet_Click()
Dim X As Integer
Dim Y As Integer
'we ask for things to customize
X = InputBox("How many cells horizontally? Now it is: " & Across)
Y = InputBox("How many cells vertically? Now it is: " & Up)
cz = InputBox("How big should a cell be (in pixels) Now it is: " & cz)

'we redim that to fit the user input
ReDim pos(X, Y)
'we cant forget about these two
Up = UBound(pos, 2)
Across = UBound(pos, 1)

'and we reinitialize all thing
ReInitialize
End Sub

Private Sub mShowLabels_Click()
If mShowLabels.Checked = False Then
MsgBox "If you have a very big map, this will slow the computer down MASSIVELY." & vbCrLf & _
        "I would not enable it, since it loads a lable for each square... only for small maps"
ShowLabels = True
mShowLabels.Checked = True
Else
ShowLabels = False
mShowLabels.Checked = False
End If
End Sub
