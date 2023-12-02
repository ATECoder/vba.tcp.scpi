VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Factory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   A public creatable singleton class for creating instances of this workbook
'''             classes to share with external workbooks. </summary>
''' <remarks>   Dependencies: DutReadingEventArgs.cls, K2700.cls, RouteSyste.cls, ScpiSystem.cls. <remarks>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

''' <summary>   Instantiates the <see cref="K2700"/> object. </summary>
''' <returns>   [<see cref="K2700"/>]. </returns>
Public Function NewK2700() As K2700
    Set NewK2700 = New K2700
End Function

''' <summary>   Gets an instance of the <see cref="DutReadingEventArgs"/> class. </summary>
''' <value>   [<see cref="DutReadingEventArgs"/>]. </returns>
Public Property Get NewDutReadingEventArgs() As DutReadingEventArgs
    Set NewDutReadingEventArgs = New DutReadingEventArgs
End Property

''' <summary>   Instantiates the <see cref="K2700ViewModel"/> object. </summary>
''' <returns>   [<see cref="K2700ViewModel"/>]. </returns>
Public Function NewK2700ViewModel() As K2700ViewModel
    Set NewK2700ViewModel = New K2700ViewModel
End Function

''' <summary>   Instantiates the <see cref="cc_isr_Tcp_Scpi.MeasureMode"/> object. </summary>
''' <returns>   [<see cref="cc_isr_Tcp_Scpi.MeasureMode"/>]. </returns>
Public Function NewMeasureMode() As cc_isr_Tcp_Scpi.MeasureMode
    Set NewMeasureMode = New cc_isr_Tcp_Scpi.MeasureMode
End Function

''' <summary>   Instantiates the <see cref="ScpiSystem"/> object. </summary>
''' <returns>   [<see cref="ScpiSystem"/>]. </returns>
Public Function NewScpiSystem() As ScpiSystem
    Set NewScpiSystem = New ScpiSystem
End Function

''' <summary>   Instantiates the <see cref="RouteSystem"/> object. </summary>
''' <returns>   [<see cref="RouteSystem"/>]. </returns>
Public Function NewRouteSystem() As RouteSystem
    Set NewRouteSystem = New RouteSystem
End Function

''' <summary>   Instantiates a new <see cref="StateChangedEventArgs"/> object. </summary>
''' <returns>   [<see cref="StateChangedEventArgs"/>]. </returns>
Public Function NewStateChangedEventArgs() As StateChangedEventArgs
    Set NewStateChangedEventArgs = New StateChangedEventArgs
End Function

