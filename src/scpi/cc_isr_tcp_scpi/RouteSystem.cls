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

Private Type this_
    Session As cc_isr_Ieee488.TcpSession
    InstrumentFamilyCards As Collection
    InstalledCards As Collection
    ChannelCount As Integer
    ScanListsFunction As String
End Type

Private This As this_

''' <summary>   Returns a reference to the <see cref="RouteSystem"/>. </summary>
''' <param name="a_session">    [<see cref="<see cref="cc_isr_Ieee488.TcpSession"/>] A connected session. </param>
''' <value>   [<see cref="RouteSystem"/>]. </returns>
Public Function Initialize(ByVal a_session As cc_isr_Ieee488.TcpSession) As RouteSystem
    Set This.Session = a_session
    Set This.InstrumentFamilyCards = New Collection
    Set This.InstalledCards = New Collection
    Set Initialize = Me
End Function

''' <summary>   Returns the connection state of the client. </summary>
''' <returns>   True if the <see cref="cc_isr_Ieee488.TcpSession"/>  is connected; otherwise, False. </returns>
Public Property Get Connected() As Boolean
    
    If This.Session Is Nothing Then
        Connected = False
    Else
        Connected = This.Session.Connected
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
        Set p_multiplexerCard = Nothing
        On Error Resume Next
        Set p_multiplexerCard = This.InstrumentFamilyCards.Item(p_cardName)
        On Error GoTo 0
        
        If Not p_multiplexerCard Is Nothing Then
            p_multiplexerCard.SlotNumber = p_slotNumber
            p_multiplexerCard.DeviceFirstChannel = p_firstChannelNumber
            p_cards.Add p_multiplexerCard.Clone(), VBA.CStr(p_slotNumber)
            p_firstChannelNumber = p_firstChannelNumber + p_multiplexerCard.FunctionalCapacity
            This.ChannelCount = This.ChannelCount + p_multiplexerCard.FunctionalCapacity
        End If
    
    Wend
    
    Set This.InstalledCards = p_cards
        
    PopulateCards = This.InstalledCards.Count
    
End Function

''' <summary>   Gets the card name. </summary>
''' <returns>   [String]. </returns>
Public Function CardName(ByVal a_slotNumber As Integer) As String
    
    If This.InstalledCards.Count >= a_slotNumber Then
        CardName = Me.InstalledCard(a_slotNumber).Name
    Else
        CardName = VBA.vbNullString
    End If

End Function

''' <summary>   Gets the card function scan list. </summary>
''' <returns>   [String]. </returns>
Public Function FunctionScanList(ByVal a_slotNumber As Integer) As String
        
    If This.InstalledCards.Count >= a_slotNumber Then
        FunctionScanList = Me.InstalledCard(a_slotNumber).FunctionScanList
    Else
        FunctionScanList = VBA.vbNullString
    End If

End Function

''' <summary>   Gets the options value and sets the cards. </summary>
''' <para name="a_deviceOptions">   [String] The device option fron the *OPT? query. </param>
''' <returns>   [Integer] the number of installed cards. </returns>
Public Function Read7700Cards(ByVal a_deviceOptions As String) As Integer

    Dim p_7700CardCount As Integer
    p_7700CardCount = Me.Populate7700Cards()
    
    Dim p_installedCardsCount As Integer
    If p_7700CardCount > 0 Then
    
        p_installedCardsCount = Me.PopulateCards(a_deviceOptions)
        
    End If
    Read7700Cards = p_installedCardsCount
    
End Function

''' <summary>   Gets the scan list function as defined in <see cref="DefineFunctionScanLists"/>. </summary>
''' <value>   [String]. </value>
Public Property Get ScanListFunction() As String
    ScanListFunction = This.ScanListsFunction
End Property

''' <summary>   Sets the complete scan lists for measurements on all cards. </summary>
''' <param name="a_function">    The measurement function. </param>
Public Sub DefineFunctionScanLists(ByVal a_function As String)

    If VBA.Len(a_function) > 0 And a_function <> This.ScanListsFunction Then
    
        Me.BuildFunctionScanLists a_function
        
        Dim p_multiplexerCard As MultiplexerCard
        For Each p_multiplexerCard In Me.InstalledCards
            This.Session.WriteLine p_multiplexerCard.FunctionScanList
        Next
    
        This.ScanListsFunction = a_function
    
    End If

End Sub

''' <summary>   Gets the total channel count for all cards. </summary>
''' <value>   [Integer]. </value>
Public Property Get ChannelCount() As Integer
    ChannelCount = This.ChannelCount
End Property

''' <summary>   Updates the functional capacity of the cards and their starting and ending
'''             channels to match any change in channel capacity.. </summary>
Public Sub UpdateFunctionalCapacities()

    Dim p_slotNumber As Integer: p_slotNumber = 0
    Dim p_firstChannelNumber As Integer
    p_firstChannelNumber = 1
    This.ChannelCount = 0

    Dim p_multiplexerCard As MultiplexerCard
    For Each p_multiplexerCard In This.InstalledCards
        p_slotNumber = p_slotNumber + 1
        If Not p_multiplexerCard Is Nothing Then
            p_multiplexerCard.SlotNumber = p_slotNumber
            p_multiplexerCard.DeviceFirstChannel = p_firstChannelNumber
            p_firstChannelNumber = p_firstChannelNumber + p_multiplexerCard.FunctionalCapacity
            This.ChannelCount = This.ChannelCount + p_multiplexerCard.FunctionalCapacity
        End If
    Next
End Sub

''' <summary>   Sets the full capacity function scan lists for on all cards. </summary>
''' <para name="a_senseFunctionName">   [String] The sense function. </param>
Public Sub BuildFunctionScanLists(ByVal a_senseFunctionName As String)
    Dim p_multiplexerCard As MultiplexerCard
    For Each p_multiplexerCard In This.InstalledCards
        p_multiplexerCard.BuildFunctionScanList a_senseFunctionName
    Next
    ' update the functional capacities based on the modified
    ' card sense function.
    Me.UpdateFunctionalCapacities
End Sub

''' <summary>   Returns the scan card for the specified device-level channel number. </summary>
''' <param name="a_deviceChannelNumber">   [Integer] The device-level channel number. </param>
''' <returns>  [<see cref="MultiplexerCard"/>]. </returns>
Public Function SelectMultiplexerCard(ByVal a_deviceChannelNumber As Integer) As MultiplexerCard
    Dim p_multiplexerCard As MultiplexerCard
    For Each p_multiplexerCard In This.InstalledCards
        If p_multiplexerCard.DeviceFirstChannel <= a_deviceChannelNumber And _
            p_multiplexerCard.DeviceLastChannel >= a_deviceChannelNumber Then
            Exit For
        End If
    Next
    If p_multiplexerCard Is Nothing Then
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidArgumentError, _
            ThisWorkbook.VBProject.Name & "." & VBA.Information.TypeName(Me) & ".SelectMultiplexerCard", _
            "Argument: 'a_deviceChannelNumber' = " & VBA.CStr(a_deviceChannelNumber) & " is out of range for the installed multiplexer cards in this instrument."
    End If
    Set SelectMultiplexerCard = p_multiplexerCard
End Function

''' <summary>   Returns the route multiple close command,
'''             e.g., <c>:ROUT:MULT:CLOS (@124,125)</c>. for the multiplexer card
'''             of the provided device-level channel number. </summary>
''' <param name="a_deviceChannelNumber">   The device-level channel number. </param>
''' <returns>  The channel list to close. </returns>
Public Function GetRouteMultipleCloseCommand(ByVal a_deviceChannelNumber As Integer) As String

    GetRouteMultipleCloseCommand = Me.SelectMultiplexerCard(a_deviceChannelNumber).RouteMultipleCloseCommand

End Function

''' <summary>   Returns the route close command,
'''             e.g., <c>:ROUT:CLOS (@124)</c>. for the multiplexer card
'''             of the provided device-level channel number. </summary>
''' <param name="a_deviceChannelNumber">   [Integer] The device-level channel number. </param>
''' <returns>  The channel list to close. </returns>
Public Function GetRouteCloseCommand(ByVal a_deviceChannelNumber As Integer) As String

    GetRouteCloseCommand = _
        Me.SelectMultiplexerCard(a_deviceChannelNumber).BuildChannelRouteCloseCommand(a_deviceChannelNumber)

End Function

''' <summary>   Returns the channel scan list, e.g., <c>(@113)</c> for a specific
'''             device-level channel number. </summary>
''' <param name="a_deviceChannelNumber">   The device-level channel number. </param>
''' <returns>  The scan list. </returns>
Public Function BuildChannelScanList(ByVal a_deviceChannelNumber As Integer) As String

    BuildChannelScanList = _
        Me.SelectMultiplexerCard(a_deviceChannelNumber).BuildChannelScanList(a_deviceChannelNumber)
    
End Function

