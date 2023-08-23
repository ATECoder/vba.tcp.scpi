Attribute VB_Name = "RouteSystemTests"
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Route System Tests. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

''' <summary>   Unit test. Asserts populating the multimplexer card 7700 cards. </summary>
''' <returns>   An <see cref="Assert"/>   instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function Test7700CardsShouldBePopulated() As Assert

    Dim p_outcome As Assert

    Dim p_routeSystem As cc_isr_Tcp_Scpi.RouteSystem
    Set p_routeSystem = cc_isr_Tcp_Scpi.Factory.NewRouteSystem.Initialize(cc_isr_Ieee488.Factory.NewViSession())
    
    Set p_outcome = Assert.IsNotNothing(p_routeSystem, TypeName(p_routeSystem) & " should be instantiated")
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.IsNotNothing(p_routeSystem.InstrumentFamilyCards, "Instrument family collection should be instantiated")
    End If
    
    
End Function













