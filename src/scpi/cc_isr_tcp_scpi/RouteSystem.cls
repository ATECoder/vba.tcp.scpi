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
    ' reference to the VI session
    Session As cc_isr_Ieee488.VISession
        InstrumentFamilyCards As Collection
        InstalledCards As Collection
End Type

Private This As ThisData

''' <summary>   Returns a reference to the <see cref="RouteSystem"/>. </summary>
''' <param name="a_session">    [<see cref="<see cref="cc_isr_IEEE488.VISession"/>"/>] A connected session. </param>
''' <value>   [<see cref="RouteSystem"/>]. </returns>
Public Function Initialize(ByVal a_session As cc_isr_Ieee488.VISession) As RouteSystem
    Set This.Session = a_session
    Set This.InstrumentFamilyCards = New Collection
    Set This.MultiplexerCards = New Collection
    Set Initialize = Me
End Function

''' <summary>   Returns the connection state of the client. </summary>
''' <returns>   True if the <see cref="cc_isr_IEEE488.VISession"/>  is connected; otherwise, False. </returns>
Public Property Get Connected() As Boolean
    
    If This.Session Is Nothing Then
        Connected = False
    Else
        Connected = This.Session.Connected
    End If

End Property

''' <summary>   Returns a reference to the <see cref="cc_isr_IEEE488.VISession"/>. </summary>
''' <value>   [<see cref="cc_isr_IEEE488.VISession"/>]. </returns>
Public Property Get Session() As cc_isr_Ieee488.VISession
    Set Session = This.Session
End Property

'''  <summary>   Gets or sets the dictionary of cards that are available for this
'''              instrument family as defined by the <see cref="Ieee488VI"/>
'''              keyed by the card name. </summary>
'''  <value> A dictionary of cards. </value>
Public Property Get InstrumentFamilyCards() As Collection
    Set InstrumentFamilyCards = This.InstrumentFamilyCards
End Property

'''  <summary>   Populates the 7700 cards family dictionary. </summary>
'''  <remarks>   2023-08-17. </remarks>
Public Sub Populate7700Cards()

        this.InstrumentFamilyCards.Clear()
        Dim p_name As String
        Dim p_capacity As Integer
        p_name = "7700": p_capacity = 20: this.InstrumentFamilyCards.Add new MultiplexerCard( p_name, p_capacity ), p_name
        p_name = "7702": p_capacity = 20: this.InstrumentFamilyCards.Add new MultiplexerCard( p_name, p_capacity ), p_name
        p_name = "7708": p_capacity = 40: this.InstrumentFamilyCards.Add new MultiplexerCard( p_name, p_capacity ), p_name
        p_name = "7710": p_capacity = 20: this.InstrumentFamilyCards.Add new MultiplexerCard( p_name, p_capacity ), p_name

End Sub

'''  <summary>   Gets or sets the installed cards for this instrument instance
'''              keyed by the card slot number. </summary>
'''  <value> The installed cards. </value>
Public Property Get InstalledCards() As Collection
    Set InstalledCards = This.InstalledCards
End Property

'''  <summary>   Populates the cards using the <see cref="Ieee488.Syntax.OptionsQueryCommand"/> reply. </summary>
'''  <remarks>   2023-08-17. </remarks>
Public Sub PopulateCards(ByVal a_cards As String)

        this.InstalledCards.Clear()
        
        Dim slotNumber As Integer: slotNumber = 0
        While (a_cards <> VBA.vbnullstring)
        
                slotNumber += 1
                dim p_cardName as String = cc_isr_Core.StringExtensions.Pop( a_cards, "," )
                If cc_isr_Core.CollectionExtensions.ContainsKey(This.InstrumentFamilyCards, p_cardName) Then
                        Dim multiplexerCard As MultiplexerCard: multiplexerCard = this.InstrumentFamilyCards[p_cardName]
                        multiplexerCard.slotNumber = slotNumber
                        this.InstalledCards.Add( multiplexerCard.Clone(), slotNumber )
                End If
        
        Wend
        
End Sub

'''  <summary>   Sets and applies the full capacity scan lists for resistance measurements on all cards. </summary>
Public Sub DefineResistanceScanLists()

        Dim p_multiplexerCard As multiplexerCard
        For Each p_multiplexerCard In This.InstalledCards
        
                _ = p_multiplexerCard.Value.BuildDefaultScanList( "RES" )
                _ = this.Session.WriteLine( p_multiplexerCard.Value.ScanList )
        Next
End Sub

'''  <summary>   Returns the route command for measuring a specific channel. </summary>
'''  <param name="channelNumber">   The channel number. </param>
'''  <returns>  The channel list to close. </returns>
Public Function BuildChannelList(ByVal a_channelNumber As Integer) As String

        Dim p_routeCommand As String: p_routeCommand = VBA.vbnullstring

        dim p_cumulativeCapacity as integer = 0
        Dim p_multiplexerCard As multiplexerCard
        For Each p_multiplexerCard In This.InstalledCards
                p_cumulativeCapacity = p_cumulativeCapacity + p_multiplexerCard.Value.Capacity
                If a_channelNumber <= p_cumulativeCapacity Then
                        p_routeCommand = p_multiplexerCard.Value.RouteCommand
        Next
        BuildChannelList = p_routeCommand

End Function

'''  <summary>   Returns the scan list for a specific channel number. </summary>
'''  <param name="channelNumber">   The channel number. </param>
'''  <returns>  The scan list. </returns>
Public Function BuildScanList(ByVal a_channelNumber As Integer) As String
{
        Dim scanCardChannel As String: scanCardChannel = VBA.vbnullstring

        Dim p_previousCumulativeCapacity As Integer: p_previousCumulativeCapacity = 0
        Dim p_cumulativeCapacity As Integer: p_cumulativeCapacity = 0
        Dim p_multiplexerCard As multiplexerCard
        For Each p_multiplexerCard In This.InstalledCards
        {
                p_cumulativeCapacity = p_cumulativeCapacity + p_multiplexerCard.Value.Capacity
                if ( a_channelNumber <= p_cumulativeCapacity )
                {
                        dim p_scanChannel as string : p_scanChannel = "0" & cstr(channelNumber - previousCumulativeCapacity}
                        p_scanCardChannel = CStr(p_multiplexerCard.Value.slotNumber) & " " & cc_isr_Core.StringExtensions.Right(scanChannel, 2)
                }
                previousCumulativeCapacity = cumulativeCapacity
        }

        BuildScanList = "(@" & scanCardChannel & ")"
}

