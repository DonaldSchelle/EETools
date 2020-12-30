Imports ExcelDna.Integration
Imports ExcelDna.IntelliSense

Module Globals
    Friend Application As Object
End Module

'Startup Function
Class StartUp
    Implements IExcelAddIn
    Public Sub Start() Implements IExcelAddIn.AutoOpen
        'Install the Intellisense server when add-in is loaded
        Application = ExcelDna.Integration.ExcelDnaUtil.Application
        IntelliSenseServer.Install()
    End Sub
    Public Sub Close() Implements IExcelAddIn.AutoClose
        'Fires when addin is removed from the addins list but not when excel closes
        ' - this is to avoid issues caused by the Excel option to cancel out of the close after the event has fired. 
        IntelliSenseServer.Uninstall()
    End Sub
End Class
