Module modStartup
    Public objAddOn As clsAddOn

    Public Sub Main()
        Try
            objAddOn = New clsAddOn
            objAddOn.Intialize()
            System.Windows.Forms.Application.Run()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
End Module
