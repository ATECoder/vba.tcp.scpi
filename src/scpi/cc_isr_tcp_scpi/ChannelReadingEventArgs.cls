VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ChannelReadingEventArgs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Channel reading event argument. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private Type this_
    ChannelNumber As Integer
    Reading As String
End Type

Private This As this_

''' <summary>   Initializes and returns the event arguments. </summary>
''' <para name="a_channelNumber">   [Integer] The channel number. </param>
''' <para name="a_reading">         [String] The reading. </param>
''' <value>   [<see cref="ChannelReadingEventArgs"/>]. </returns>
Public Function Initialize(ByVal a_channelNumber As Integer, ByVal a_reading As String) As ChannelReadingEventArgs
    This.ChannelNumber = a_channelNumber
    This.Reading = a_reading
    Set Initialize = Me
End Function

''' <summary>   Gets the channel number. </summary>
''' <value>   [Integer]. </value>
Public Property Get ChannelNumber() As Integer
    ChannelNumber = This.ChannelNumber
End Property

''' <summary>   Gets the reading. </summary>
''' <value>   [String]. </value>
Public Property Get Reading() As String
    Reading = This.Reading
End Property

