Imports GetFileAndDirListLib

Public Class ConsoleDemo
    Public GetFileAndDirListApp As GetFileAndDirListApp
    Public RootDirPath As String

    Public Sub Main()
        Do While True
            Console.WriteLine("*******************")
            Console.WriteLine("Main menu")
            Console.WriteLine("*******************")
            Console.WriteLine("Press Q to Exit")
            Console.WriteLine("Press A to Scan Folders Dir and Files")
            '            Console.WriteLine("Press B to Test Linux")
            Console.WriteLine("*******************")
            Select Case Console.ReadKey().Key
                Case ConsoleKey.Q
                    Exit Do
                Case ConsoleKey.A
                    Console.WriteLine("*******************")
                    Console.WriteLine("Scan Folders Dir and Files")
                    Console.WriteLine("*******************")
                    Console.WriteLine("Enter the starting folder for the scan")
                    Me.RootDirPath = Console.ReadLine
                    Me.GetFileAndDirListApp = New GetFileAndDirListApp(Me.RootDirPath)
                    Console.WriteLine("RootDirPath=" & Me.RootDirPath)
                    '                    Console.WriteLine("LinuxTest=" & Me.GetFileAndDirListApp.LinuxTest)
                    If Me.GetFileAndDirListApp.LastErr <> "" Then
                        Console.WriteLine(Me.GetFileAndDirListApp.LastErr)
                    Else
                        Console.WriteLine("Start" & Me.GetFileAndDirListApp.DirListPath)
                        Me.GetFileAndDirListApp.Start()
                        If Me.GetFileAndDirListApp.LastErr <> "" Then
                            Console.WriteLine(Me.GetFileAndDirListApp.LastErr)
                        Else
                            Console.WriteLine("Scanning succeeded, please view DirList.txt and FileList.txt.")
                        End If
                    End If
                    'Case ConsoleKey.B
                    '    Me.GetFileAndDirListApp = New GetFileAndDirListApp
                    '    Console.WriteLine("LinuxTest=" & Me.GetFileAndDirListApp.LinuxTest)
            End Select
        Loop
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

End Class
