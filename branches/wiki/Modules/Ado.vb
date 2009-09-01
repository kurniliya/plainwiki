Namespace Openwiki
    Module Ado
        ' Constants used by MS ADO.DB

        '---- CursorTypeEnum Values ----
        Public Const adOpenForwardOnly As Integer = 0
        Public Const adOpenKeyset As Integer = 1
        Public Const adOpenDynamic As Integer = 2
        Public Const adOpenStatic As Integer = 3

        '---- LockTypeEnum Values ----
        Public Const adLockReadOnly As Integer = 1
        Public Const adLockPessimistic As Integer = 2
        Public Const adLockOptimistic As Integer = 3
        Public Const adLockBatchOptimistic As Integer = 4

        '---- CursorLocationEnum Values ----
        Public Const adUseServer As Integer = 2
        Public Const adUseClient As Integer = 3

        '---- SearchDirection Values ----
        Public Const adSearchForward As Integer = 1
        Public Const adSearchBackward As Integer = -1

        '---- CommandTypeEnum Values ----
        Public Const adCmdUnknown As Integer = &H8
        Public Const adCmdText As Integer = &H1
        Public Const adCmdTable As Integer = &H2
        Public Const adCmdStoredProc As Integer = &H4

    End Module
End Namespace