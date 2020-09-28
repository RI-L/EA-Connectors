Option Explicit

!INC Local Scripts.EAConstants-VBScript
!INC Connectors.Connector
!INC Connectors.Connector_Utils

''' ------------------------
''' TESTING CONNECTOR HELPER
''' ------------------------

Private Sub Module_Initialize()
	''' Ensure initialization of the variable as to prepare 
	''' for the assigment check in the ConAPI() function
	Set m_conapi = Nothing
End Sub

Private Sub Module_Terminate()
	''' Call this explicitly to dispose of the object.
	Set m_conapi = Nothing
End Sub

Private Sub PrintStr(msg)
	Session.Output msg
End SUb

Private Sub PrintWrappedConnectors()

	''' Print repository
	PrintStr "----------------------------"
	PrintStr " Testing Connection Wrapper "
	PrintStr "----------------------------"
	PrintStr " "
	
	Dim Package as EA.Package
	Set Package = Repository.GetTreeSelectedPackage()

'	''' PACKAGES
'	Dim pkg as EA.Package
'	For Each pkg in Package.Packages
	
	''' CLASSES
	Dim e as EA.Element
	For Each e in Package.Elements
		PrintStr "Class : " & e.Name

		Dim cw
		Set cw = ConApi()						''' Local variables are always faster

		''' CONNECTORS
		Dim c As EA.Connector
		For Each c in e.Connectors
			PrintStr "Connector 1			: " & cw.Wrap(e, c).AutoName(True).Name()
			
			''' To reset an aggregation you must reset both ends!
			PrintStr "Connector 2			: " & cw.SetNearEndAggregation(EA_none).Name()
			PrintStr "Connector 3			: " & cw.SetOtherEndAggregation(EA_none).Name()
			
			''' Set either shared or composite in either end (EA_none, EA_shared, EA_composite)
			''' The other end is autmatically set to EA_none :
			PrintStr "Connector 4			: " & cw.SetNearEndAggregation(EA_composite).Name()
			''' or
			PrintStr "Connector 5			: " & cw.SetOtherEndAggregation(EA_shared).Name()
			
			''' Or set both the AutoName AND the Aggregation AND get the Name at once! :
			PrintStr "Connector 6			: " & cw.Wrap(e, c).AutoName(True).SetOtherEndAggregation(EA_shared).Name()
			
			
			
			PrintStr "IsOwnedByNearEnd: " & cw.IsOwnedByNearEnd()	''' ***
			
			PrintStr "Direction		    : " & cw.Direction()
			
			PrintStr "Type				: " & cw.Type_()
			PrintStr "IsAutoName			: " & cw.IsAutoName()
			
			PrintStr "IsAssociationKind: " & cw.IsAssociationKind
			PrintStr "IsGeneralization	: " & cw.IsGeneralization
			PrintStr "IsAssociation		: " & cw.IsAssociation
			PrintStr "IsAggregation		: " & cw.IsAggregation
			PrintStr "IsComposition		: " & cw.IsComposition
			PrintStr "IsRealisation		: " & cw.IsRealisation
			PrintStr "IsDependency		: " & cw.IsDependency
			
			PrintStr "IsNearEndAggregate: " & cw.IsNearEndAggregate

			''' Hmm....
			PrintStr "IsAggregate		 : " & cw.IsAggregate
			PrintStr "IsComposite		 : " & cw.IsComposite
			PrintStr "IsAggregationKind  : " & cw.IsAggregationKind			

			''' Roles
					
			PrintStr "nearEnd.EaEnd		: " & cw.nearEnd.End_
			PrintStr "otherEnd.EaEnd		: " & cw.otherEnd.End_

			PrintStr "ne.Aggregation		: " & cw.nearEnd.Aggregation
			PrintStr "ot.Aggregation		: " & cw.otherEnd.Aggregation
			
			PrintStr "ne.IsAggregationKind: " & cw.nearEnd.IsAggregationKind
			PrintStr "ot.IsAggregationKind: " & cw.otherEnd.IsAggregationKind

			PrintStr "ne.IsAggregate		: " & cw.nearEnd.IsAggregate
			PrintStr "ot.IsAggregate		: " & cw.otherEnd.IsAggregate

			PrintStr "ne.IsComposite		: " & cw.nearEnd.IsComposite
			PrintStr "ot.IsComposite		: " & cw.otherEnd.IsComposite

			PrintStr "ne.EndClassName		: " & cw.nearEnd.EndClassName
			PrintStr "ot.EndClassName		: " & cw.otherEnd.EndClassName
			
			PrintStr "ne.IsAssociationKind: " & cw.nearEnd.IsAssociationKind
			PrintStr "ot.IsAssociationKind: " & cw.otherEnd.IsAssociationKind

			PrintStr "ne.IsExplicitlyNavigable: " & cw.nearEnd.IsExplicitlyNavigable
			PrintStr "ot.IsExplicitlyNavigable: " & cw.otherEnd.IsExplicitlyNavigable
			
			PrintStr "ne.IsNavigable		: " & cw.nearEnd.IsNavigable
			PrintStr "ot.IsNavigable		: " & cw.otherEnd.IsNavigable

			PrintStr "ne.Navigability		: " & cw.nearEnd.Navigability
			PrintStr "ot.Navigability		: " & cw.otherEnd.Navigability


			PrintStr "ne.IsClientEnd		: " & cw.nearEnd.IsClientEnd
			PrintStr "ot.IsClientEnd		: " & cw.otherEnd.IsClientEnd
			PrintStr "ne.IsSupplierEnd	: " & cw.nearEnd.IsSupplierEnd
			PrintStr "ot.IsSupplierEnd	: " & cw.otherEnd.IsSupplierEnd

			PrintStr "ne.IsNearEnd		: " & cw.nearEnd.IsNearEnd
			PrintStr "ot.IsOtherEnd		: " & cw.otherEnd.IsOtherEnd
			PrintStr "ne.IsOtherEnd		: " & cw.nearEnd.IsOtherEnd
			PrintStr "ot.IsNearEnd		: " & cw.otherEnd.IsNearEnd

			PrintStr "ne.Visibility		: " & cw.nearEnd.Visibility
			PrintStr "ot.Visibility		: " & cw.otherEnd.Visibility

			
			
		Next ''' Connector
		PrintStr "---------------"
	Next ''' Class
'	Next ''' Pkg
		

'	ConApi.StatsStop
'	if ConApi.HasStats then
'		PrintStr "Statistics"
'		PrintStr "-----------------------------------------------"
'		PrintStr "Hits         (Wraps): " & cw.StatsWrapCount()
'		PrintStr "Hits                : " & cw.StatsCount()
'		PrintStr "Hits           (Acc): " & cw.StatsCountAcc()
'		PrintStr "Hits Per Second     : " & Round( cw.StatsHitsPerSecond(), 3)
'		PrintStr "Hits Per Second(Acc): " & Round( cw.StatsHitsPerSecondAcc(), 3)
'		PrintStr "-----------------------------------------------"
'		PrintStr "Duration            : " & Minute( cw.StatsDuration()) & ":" & Round( Second(cw.StatsDuration()), 3)
'		PrintStr "Duration       (Acc): " & Minute( cw.StatsDurationAcc()) & ":" & Round( Second(cw.StatsDurationAcc()), 3)
'		PrintStr "Time Per Hit        : " & Round( Second(cw.StatsTimePerHits()), 3) & " sec"
'		PrintStr "Time Per Hit   (Acc): " & Round( Second(cw.StatsTimePerHitsAcc()), 3) & " sec"
'		PrintStr "-----------------------------------------------"
'		PrintStr " "
'	End If
	Set cw = Nothing
	Set m_conapi = Nothing
End Sub



''' MAIN

Sub Main()
	Repository.EnsureOutputVisible( "Script" )
	Module_Initialize()
	
	PrintStr "--++:::: START: " & Date() & " -- " & Time() & " ::::++--"
	PrintWrappedConnectors()
	PrintStr "--++:::: STOP: " & Date() & " -- " & Time() & " ::::++--"	
	
	Module_Terminate()	
End Sub

Main 