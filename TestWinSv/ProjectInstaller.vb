Imports System.ComponentModel
Imports System.Configuration.Install

Public Class ProjectInstaller

    Public Sub New()
        MyBase.New()

        'This call is required by the Component Designer.
        InitializeComponent()

        'Add initialization code after the call to InitializeComponent

    End Sub

    Private Sub RequestService_AfterInstall(sender As Object, e As InstallEventArgs) Handles RequestService.AfterInstall
        Using serviceController As New System.ServiceProcess.ServiceController(RequestService.ServiceName)
            serviceController.Start()
        End Using
    End Sub
End Class
