# PigCmdLibDemo

Demo console program for PigCmdLib.

```
VB.net
Module Program
    Sub Main(args As String())
        Dim oConsoleDemo As New PigCmdLibDemo.ConsoleDemo
        oConsoleDemo.Main()
    End Sub
End Module

C#
class Program
{
    static void Main(string[] args)
    {
        PigCmdLibDemo.ConsoleDemo oConsoleDemo = new PigCmdLibDemo.ConsoleDemo() ;
        oConsoleDemo.Main();
    }
}

```