VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ErrTracer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements cc_isr_Core.IErrTracer

Private Sub IErrTracer_TraceError()
    K2700Sheet.LastErrorMessage = cc_isr_Core.ThisWorkbook.BuildStandardErrorMessage
End Sub

