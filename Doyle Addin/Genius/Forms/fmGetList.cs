#error Cannot convert CompilationUnitSyntax - see comment for details
/* Cannot convert CompilationUnitSyntax, System.NullReferenceException: Object reference not set to an instance of an object.
   at ICSharpCode.CodeConverter.CSharp.DeclarationNodeVisitor.ShouldBeNestedInRootNamespace(StatementSyntax vbStatement, String rootNamespace)
   at System.Linq.Lookup`2.Create[TSource](IEnumerable`1 source, Func`2 keySelector, Func`2 elementSelector, IEqualityComparer`1 comparer)
   at ICSharpCode.CodeConverter.CSharp.DeclarationNodeVisitor.PrependRootNamespace(IReadOnlyCollection`1 membersConversions, String rootNamespace)
   at ICSharpCode.CodeConverter.CSharp.DeclarationNodeVisitor.<VisitCompilationUnit>d__31.MoveNext()
--- End of stack trace from previous location where exception was thrown ---
   at System.Runtime.ExceptionServices.ExceptionDispatchInfo.Throw()
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingVisitorWrapper.<ConvertHandledAsync>d__12`1.MoveNext()

Input:
VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fmGetList 
   Caption         =   "List Entry"
   ClientHeight    =   6975
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3525
   OleObjectBlob   =   "fmGetList.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fmGetList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'Event CheckOut(Cancel As Long)

Private bg As String
Private rt As String

Public Function AskUser( _
    Optional Using As String = "" _
) As String
    Me.bg = Using          'save initial text
    txIn.Value = Me.bg     'initialize text box
    Show vbModal        'and wait...
    AskUser = Me.rt        'return final result
End Function

Private Sub CheckOut(NoChg As Long)
    Dim ck As VbMsgBoxResult
    
    If NoChg = 0 Then
        ck = Global.Microsoft.VisualBasic.Interaction.MsgBox( _
            "Use this List?", _
            Global.Microsoft.VisualBasic.Constants.vbYesNo + Global.Microsoft.VisualBasic.Constants.vbQuestion, _
            "Confirm" _
        )
        If ck = Global.Microsoft.VisualBasic.Constants.vbYes Then Me.rt = txIn.Value
    Else
        ck = Global.Microsoft.VisualBasic.Interaction.MsgBox( _
            "Cancel this Entry?", _
            Global.Microsoft.VisualBasic.Constants.vbYesNo + Global.Microsoft.VisualBasic.Constants.vbQuestion, _
            "Cancel" _
        )
        If ck = Global.Microsoft.VisualBasic.Constants.vbYes Then Me.rt = Me.bg
    End If
    
    If ck = Global.Microsoft.VisualBasic.Constants.vbYes Then Me.Hide
End Sub

Private Sub cmdCancel_Click()
    Me.CheckOut 1 'no change
End Sub

Private Sub cmdOk_Click()
    Me.CheckOut 0 'commit changes
End Sub

Private Sub UserForm_QueryClose( _
    Cancel As Integer, _
    CloseMode As Integer _
)
    Cancel = 1
    Me.CheckOut 1 'no change
End Sub

Private Sub UserForm_Initialize()
    '''
End Sub

Private Sub UserForm_Terminate()
    '''
End Sub

 */