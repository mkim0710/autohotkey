#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

if not A_IsAdmin
{
   Run *RunAs "%A_ScriptFullPath%"  ; Requires v1.0.92.01+
   ExitApp
}

CoordMode, Mouse, Screen

MoveClickBack(x, y)
{
   local start_x, start_y
   mousegetpos, start_x, start_y
   mouseclick, left, %x%, %y%, 1, 0
   Sleep 300                      
   mousemove, %start_x%, %start_y%, 0
}
!l::
   mousegetpos, start_x, start_y
   mouseclick, left, 5, 257, 1, 0
   Sleep 500                      
   mouseclick, left, 1630, 140, 1, 0
   Sleep 300                      
   mousemove, %start_x%, %start_y%, 0
return

!h::MoveClickBack(260, 360) 
return

!d::
   mousegetpos, start_x, start_y
   mouseclick, left, 930, 65, 1, 0
   
   MsgBox, 1, , Confirm Diagnosis? (Press Ok), 3
   mouseclick, left, 1010, 400, 1, 0
   
   MsgBox, 1, , Confirm Diagnosis? (Press Ok), 3
   mouseclick, left, 960, 920, 1, 0
   
   Sleep 800                      
   mouseclick, left, 970, 570, 1, 0
   
   MsgBox, 1, , Patient List? (Press Ok), 3
   mouseclick, left, 5, 257, 1, 0
   
   Sleep 1000                      
   mouseclick, left, 1630, 140, 1, 0
return


!f::
   mouseclick, left, 5, 257, 1, 0
   Sleep 1000                      
   mouseclick, left, 1630, 140, 1, 0
return
