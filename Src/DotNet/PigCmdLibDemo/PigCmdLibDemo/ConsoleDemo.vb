Imports PigCmdLib

Public Class ConsoleDemo
    Public CmdOrFilePath As String
    Public CmdPara As String
    Public PigCmdApp As New PigCmdApp
    Public PID As Long
    Public Line As String
    Public Sub Main()
        Do While True
            Console.WriteLine("*******************")
            Console.WriteLine("Main menu")
            Console.WriteLine("*******************")
            Console.WriteLine("Press Q to Exit")
            Console.WriteLine("Press A to HideShell")
            Console.WriteLine("*******************")
            Select Case Console.ReadKey().Key
                Case ConsoleKey.Q
                    Exit Do
                Case ConsoleKey.A
                    Console.WriteLine("*******************")
                    Console.WriteLine("HideShell")
                    Console.WriteLine("*******************")
                    Console.WriteLine("CmdFilePath=" & Me.CmdOrFilePath)
                    Me.Line = Console.ReadLine
                    If Me.Line <> "" Then
                        Me.CmdOrFilePath = Me.Line
                    End If
                    Me.PID = PigCmdApp.HideShell(Me.CmdOrFilePath)
                    Console.WriteLine("LastErr=" & Me.PigCmdApp.LastErr)
                    Console.WriteLine("PID=" & Me.PID)
            End Select
        Loop
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

End Class
