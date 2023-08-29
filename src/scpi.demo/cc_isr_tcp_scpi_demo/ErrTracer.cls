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

Private Type this_
    ViewModel As K2700ViewModel
End Type

Private This As this_

Implements cc_isr_Core.IErrTracer

Public Function Initialize(ByVal a_viewModel As K2700ViewModel)
    This.ViewModel = a_viewModel
End Function

Private Sub IErrTracer_TraceError()
    This.ViewModel.LastErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage
End Sub

