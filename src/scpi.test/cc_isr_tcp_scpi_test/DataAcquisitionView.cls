VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataAcquisitionView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>    The Data Acquisition View for linking the observer with the user interface. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private Const m_sessionTimeout As String = "B1"
Private Const m_inputAddress As String = "B2"
Private Const m_inputPort As String = "B3"
Private Const m_socketAddress As String = "C2"
Private Const m_socketId As String = "D2"
Private Const m_connected As String = "E2"

Private Const m_timerInterval As String = "C22"

Private Const m_lastMessage As String = "B4"
Private Const m_lastErrorMessage As String = "B33"

Private Const m_measurementRow As Integer = 2
Private Const m_measurementCol As Integer = 23
Private Const m_channelRow As Integer = 2
Private Const m_channelCol As Integer = 24
Private Const m_byRow As Boolean = True
Private Const m_displayIncrement As Integer = 1

' Controls

Private WithEvents AutoIncrementCheckBox As MSForms.CheckBox
Attribute AutoIncrementCheckBox.VB_VarHelpID = -1
Private WithEvents ClearReadingsButton As MSForms.CommandButton
Attribute ClearReadingsButton.VB_VarHelpID = -1

Private Type this_
    MeasurementRow As Integer
    MeasurementColumn As Integer
    ChannelRow As Integer
    ChannelColumn As Integer
    ChannelCellValue As String
    MeasurementCellValue As String

    Observer As K2700Observer

End Type

Private This As this_

' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Construction and cleanup
' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Private Sub Class_Initialize()
    Set This.Observer = K2700Observer
End Sub

Private Sub Class_Terminate()
    Set This.Observer = Nothing
End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' Data Acquisition Controls
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Handles the auto increment check box change and click events. </summary>
Private Sub AutoIncrementCheckBox_Change()
    This.Observer.AutoIncrementChannelNoEnabled = AutoIncrementCheckBox.Value
End Sub

Public Property Let AutoIncrementCheckBoxEnabled(ByVal a_value As Boolean)
    If Not AutoIncrementCheckBox Is Nothing Then
        AutoIncrementCheckBox.Enabled = a_value
    End If
End Property

Public Property Let AutoIncrementCheckBoxValue(ByVal a_value As Boolean)
    If Not AutoIncrementCheckBox Is Nothing Then
        AutoIncrementCheckBox.Value = a_value
    End If
End Property

Private Sub ClearReadingsButton_Click()
    This.Observer.OnClearReadings
End Sub

Public Property Let ClearReadingsButtonEnabled(ByVal a_value As Boolean)
    If Not ClearReadingsButton Is Nothing Then
        ClearReadingsButton.Enabled = a_value
    End If
End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' View Model Event Handlers
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Gets the Measurement Row. </summary>
''' <value>   [Integer]. </value>
Public Property Get MeasurementRow() As Integer
    MeasurementRow = This.MeasurementRow
End Property

''' <summary>   Sets the Measurement Row. </summary>
''' <param name="a_value">   [Integer] the Measurement Row. </value>
Public Property Let MeasurementRow(ByVal a_value As Integer)
    If This.MeasurementRow <> a_value Then
        This.MeasurementRow = a_value
    End If
End Property

''' <summary>   Gets the Measurement Column. </summary>
''' <value>   [Integer]. </value>
Public Property Get MeasurementColumn() As Integer
    MeasurementColumn = This.MeasurementColumn
End Property

''' <summary>   Sets the Measurement Column. </summary>
''' <param name="a_value">   [Integer] the Measurement Column. </value>
Public Property Let MeasurementColumn(ByVal a_value As Integer)
    If This.MeasurementColumn <> a_value Then
        This.MeasurementColumn = a_value
    End If
End Property

''' <summary>   Gets the Channel Row. </summary>
''' <value>   [Integer]. </value>
Public Property Get ChannelRow() As Integer
    ChannelRow = This.ChannelRow
End Property

''' <summary>   Sets the Channel Row. </summary>
''' <param name="a_value">   [Integer] the Channel Row. </value>
Public Property Let ChannelRow(ByVal a_value As Integer)
    If This.ChannelRow <> a_value Then
        This.ChannelRow = a_value
    End If
End Property

''' <summary>   Gets the Channel Column. </summary>
''' <value>   [Integer]. </value>
Public Property Get ChannelColumn() As Integer
    ChannelColumn = This.ChannelColumn
End Property

''' <summary>   Sets the Channel Column. </summary>
''' <param name="a_value">   [Integer] the Channel Column. </value>
Public Property Let ChannelColumn(ByVal a_value As Integer)
    If This.ChannelColumn <> a_value Then
        This.ChannelColumn = a_value
    End If
End Property

''' <summary>   Gets the Measurement Cell Value. </summary>
''' <value>   [String]. </value>
Public Property Get MeasurementCellValue() As String
    MeasurementCellValue = This.MeasurementCellValue
    ' MeasurementCellValue = Cells(Me.MeasurementRow, Me.MeasurementColumn).Value
End Property

''' <summary>   Sets the Measurement Cell Value. </summary>
''' <param name="a_value">   [String] the Measurement Cell Value. </value>
Public Property Let MeasurementCellValue(ByVal a_value As String)
    If This.MeasurementCellValue <> a_value Then
        This.MeasurementCellValue = a_value
        ' Cells(Me.MeasurementRow, Me.MeasurementColumn).Value = a_value
    End If
End Property
        
''' <summary>   Gets the channel cell Value. </summary>
''' <value>   [String]. </value>
Public Property Get ChannelCellValue() As String
    ChannelCellValue = This.ChannelCellValue
    ' ChannelCellValue = Cells(Me.ChannelRow, Me.ChannelColumn).Value
End Property

''' <summary>   Sets the channel cell value. </summary>
''' <param name="a_value">   [String] the channel cell value. </value>
Public Property Let ChannelCellValue(ByVal a_value As String)
    If This.ChannelCellValue <> a_value Then
        This.ChannelCellValue = a_value
        ' Cells(Me.ChannelRow, Me.ChannelColumn).Value = a_value
    End If
End Property

''' <summary>   Sets the Channel display. </summary>
''' <param name="a_channelNo">   [single] the channel number. </param>
''' <param name="a_value">       [double] the value. </param>
''' <param name="a_format">      [String] the format string for formatting the value. </param>
Public Sub SetReading(ByVal a_channelNo As Integer, ByVal a_value As Double, a_format As String)
    SetReading_ a_channelNo, a_value, m_byRow, m_displayIncrement, a_format
End Sub

''' <summary>   Sets the Channel display. </summary>
''' <param name="a_channelNo">   [single] the channel number. </param>
''' <param name="a_value">       [double] the value. </param>
''' <param name="a_byRow">       [Boolean] true if the next reading goes to the next row. </param>
''' <param name="a_increment">   [Integer] the row or column increment. </param>
''' <param name="a_format">      [String] the format string for formatting the value. </param>
Private Sub SetReading_(ByVal a_channelNo As Integer, ByVal a_value As Double, _
    ByVal a_byRow As Boolean, ByVal a_increment As Integer, a_format As String)

    Me.MeasurementRow = IIf(a_byRow, m_measurementRow + a_increment * (a_channelNo - 1), m_measurementRow)
    Me.MeasurementColumn = IIf(a_byRow, m_measurementCol, m_measurementCol + a_increment * (a_channelNo - 1))
   
    Me.ChannelRow = IIf(a_byRow, m_channelRow + a_increment * (a_channelNo - 1), m_channelRow)
    Me.ChannelColumn = IIf(a_byRow, m_channelCol, m_channelCol + a_increment * (a_channelNo - 1))
   
    Me.ChannelCellValue = VBA.CStr(a_channelNo)
    Me.MeasurementCellValue = VBA.Format$(a_value, a_format)
    
End Sub



