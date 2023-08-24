VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RouteSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   SCPI Route Subsystem Implementation. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private Type ThisData
    ViSession As cc_isr_Ieee488.ViSession
    InstrumentFamilyCards As Collection
    InstalledCards As Collection
    ChannelCount As Integer
End Type

Private This As ThisData

''' <summary>   Returns a reference to the <see cref="RouteSystem"/>. </summary>
''' <param name="a_session">    [<see cref="<see cref="cc_isr_IEEE488.VISession"/>"/>] A connected session. </param>
''' <value>   [<see cref="RouteSystem"/>]. </returns>
Public Function Initialize(ByVal a_session As cc_isr_Ieee488.ViSession) As RouteSystem
    Set This.ViSession = a_session
    Set This.InstrumentFamilyCards = New Collection
    Set This.InstalledCards = New Collection
    Set Initialize = Me
End Function

''' <summary>   Returns the connection state of the client. </summary>
''' <returns>   True if the <see cref="cc_isr_IEEE488.VISession"/>  is connected; otherwise, False. </returns>
Public Property Get Connected() As Boolean
    
    If This.ViSession Is Nothing Then
        Connected = False
    Else
        Connected = This.ViSession.Connected
    End If

End Property

''' <summary>   Gets or sets the dictionary of cards that are available for this
'''              instrument family as defined by the <see cref="Ieee488VI"/>
'''              keyed by the card name. </summary>
''' <value> A dictionary of cards. </value>
Public Property Get InstrumentFamilyCards() As Collection
    Set InstrumentFamilyCards = This.InstrumentFamilyCards
End Property

''' <summary>   Populates the 7700 cards family dictionary. </summary>
''' <remarks>   2023-08-17. </remarks>
''' <returns>   [Integer] the number of items. </returns>
Public Function Populate7700Cards() As Integer

    Dim p_cards As New Collection
    Dim p_name As String
    Dim p_capacity As Integer
    Dim p_card As MultiplexerCard
    Set p_card = New MultiplexerCard: p_cards.Add p_card.Initialize("7700", 20), p_card.Name
    Set p_card = New MultiplexerCard: p_cards.Add p_card.Initialize("7702", 40), p_card.Name
    Set p_card = New MultiplexerCard: p_cards.Add p_card.Initialize("7708", 40), p_card.Name
    Set p_card = New MultiplexerCard: p_cards.Add p_card.Initialize("7710", 20), p_card.Name
    Set This.InstrumentFamilyCards = p_cards
    Populate7700Cards = This.InstrumentFamilyCards.Count

End Function

''' <summary>   Gets an installed card. </summary>
''' <value>   [<see cref=MultiplexerCard"/>]. </value>
Public Function InstalledCard(ByVal a_slotNumber As Integer) As MultiplexerCard
    
    If This.InstalledCards.Count >= a_slotNumber Then
        Set InstalledCard = This.InstalledCards(a_slotNumber)
    Else
        Set InstalledCard = Nothing
    End If
    
End Function

''' <summary>   Gets or sets the installed cards for this instrument instance
'''              keyed by the card slot number. </summary>
''' <value>   [Collection] The installed cards. </value>
Public Property Get InstalledCards() As Collection
    Set InstalledCards = This.InstalledCards
End Property

''' <summary>   Populates the installed cards using the <see cref="Ieee488.Syntax.OptionsQueryCommand"/> reply. </summary>
''' <remarks>   2023-08-17. </remarks>
''' <value>   [Integer] the number of installed cards. </value>
Public Function PopulateCards(ByVal a_cards As String) As Integer

    Dim p_cards As New Collection
    Dim p_slotNumber As Integer: p_slotNumber = 0
    Dim p_firstChannelNumber As Integer
    p_firstChannelNumber = 1
    This.ChannelCount = 0
    While (a_cards <> VBA.vbNullString)
        
        p_slotNumber = p_slotNumber + 1
        Dim p_cardName As String
        p_cardName = cc_isr_Core.StringExtensions.Pop(a_cards, ",")
        Dim p_multiplexerCard As MultiplexerCard
        
        On Error Resume Next
        Set p_multiplexerCard = This.InstrumentFamilyCards.Item(p_cardName)
        On Error GoTo 0
        
        If Not p_multiplexerCard Is Nothing Then
            p_multiplexerCard.SlotNumber = p_slotNumber
            p_multiplexerCard.DeviceFirstChannel = p_firstChannelNumber
            p_cards.Add p_multiplexerCard.Clone(), CStr(p_slotNumber)
            p_firstChannelNumber = p_firstChannelNumber + p_multiplexerCard.Capacity
            This.ChannelCount = This.ChannelCount + p_multiplexerCard.Capacity
        End If
    
    Wend
    
    Set This.InstalledCards = p_cards
        
    PopulateCards = This.InstalledCards.Count
    
End Function

''' <summary>   Gets the total channel count for all cards. </summary>
''' <value>   [Integer]. </value>
Public Property Get ChannelCount() As Integer
    ChannelCount = This.ChannelCount
End Property

''' <summary>   Sets the full capacity function scan lists for on all cards. </summary>
Public Sub BuildFunctionScanLists(ByVal a_scpiFunction As String)

    Dim p_multiplexerCard As MultiplexerCard
    For Each p_multiplexerCard In This.InstalledCards
        p_multiplexerCard.BuildFunctionScanList a_scpiFunction
    Next
End Sub

''' <summary>   Returns the scan card for the specific channel. </summary>
''' <param name="a_channelNumber">   The channel number. </param>
''' <returns>  [<see cref="MultiplexerCard"/>]. </returns>
Public Function SelectMultiplexerCard(ByVal a_channelNumber As Integer) As MultiplexerCard
    Dim p_multiplexerCard As MultiplexerCard
    For Each p_multiplexerCard In This.InstalledCards
        If p_multiplexerCard.DeviceFirstChannel <= a_channelNumber And _
            p_multiplexerCard.DeviceLastChannel >= a_channelNumber Then
            Exit For
        End If
    Next
    If p_multiplexerCard Is Nothing Then
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidArgumentError, _
            ThisWorkbook.VBProject.Name & "." & VBA.Information.TypeName(Me) & ".SelectMultiplexerCard", _
            "Argument: 'a_ChannelNumber' = " & CStr(a_channelNumber) & " is out of range for the installed multiplexer cards in this instrument."
    End If
    Set SelectMultiplexerCard = p_multiplexerCard
End Function

''' <summary>   Returns the route multiple close command,
'''             e.g., <c>:ROUT:MULT:CLOS (@124,125)</c>. for the multiplexer card
'''             of the provided device channel number. </summary>
''' <param name="a_channelNumber">   The channel number. </param>
''' <returns>  The channel list to close. </returns>
Public Function GetRouteMultipleCloseCommand(ByVal a_channelNumber As Integer) As String

    GetRouteMultipleCloseCommand = Me.SelectMultiplexerCard(a_channelNumber).RouteMultipleCloseCommand

End Function

''' <summary>   Returns the route close command,
'''             e.g., <c>:ROUT:CLOS (@124)</c>. for the multiplexer card
'''             of the provided device channel number. </summary>
''' <param name="a_channelNumber">   The channel number. </param>
''' <returns>  The channel list to close. </returns>
Public Function GetRouteCloseCommand(ByVal a_channelNumber As Integer) As String

    GetRouteCloseCommand = _
        Me.SelectMultiplexerCard(a_channelNumber).BuildChannelRouteCloseCommand(a_channelNumber)

End Function

''' <summary>   Returns the channel scan list, e.g., <c>(@113)</c> for a specific
'''             device channel number. </summary>
''' <param name="a_channelNumber">   The device channel number. </param>
''' <returns>  The scan list. </returns>
Public Function BuildChannelScanList(ByVal a_channelNumber As Integer) As String

    BuildChannelScanList = _
        Me.SelectMultiplexerCard(a_channelNumber).BuildChannelScanList(a_channelNumber)
    
End Function

