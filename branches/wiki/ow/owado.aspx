<script language="VB" runat="Server">

' Constants used by MS ADO.DB

'---- CursorTypeEnum Values ----
Const adOpenForwardOnly As Short = 0

Const adOpenKeyset As Short = 1

Const adOpenDynamic As Short = 2

Const adOpenStatic As Short = 3


'---- LockTypeEnum Values ----
Const adLockReadOnly As Short = 1

Const adLockPessimistic As Short = 2

Const adLockOptimistic As Short = 3

Const adLockBatchOptimistic As Short = 4


'---- CursorLocationEnum Values ----
Const adUseServer As Short = 2

Const adUseClient As Short = 3


'---- SearchDirection Values ----
Const adSearchForward As Short = 1

Const adSearchBackward As Short = -1


'---- CommandTypeEnum Values ----
Const adCmdUnknown As Short = &H8s

Const adCmdText As Short = &H1s

Const adCmdTable As Short = &H2s

Const adCmdStoredProc As Short = &H4s

</script>
