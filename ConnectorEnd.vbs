'Option Explicit

!INC Local Scripts.EAConstants-VBScript
!INC TaggedValues.Wrappers_Const

''' ===========================================================================
''' WRAPPER for CONNECTOR END
''' ===========================================================================
'''
''' VERSION			: 0.9.1, 20200927 - Edited description text.
''' 				: 0.9.0, 20151201 - initial commit, 
'''
''' DESCRIPTION		: A Connector Helper wrapper intending to provide access 
'''					  to EA Connector properties and concepts with a consistent
''' 				  "user perspective" of associations rather than only the
'''					  technical terms for the EA model elements.
'''					  This means in practice for example that end roles, as seen 
'''					  from the viewpoint class, are named <CurrentClass>.nearEnd 
'''                   and the role of the far end again, as viewed from the
'''					  calling class, is named <CurrentClass>.otherEnd.
''' 				  In short: The "otherEnd" would correspond to the roleName
'''					  of the actual class member of the class at the "nearEnd."
'''
''' AUTHOR			: Rolf Lampa, RIL Partner AB, laromai@rilnet.com
'''
''' COPYRIGHT		: (C) Rolf Lampa, 2015. Free to use for commercial projects 
'''				  	  if giving proper attribution to the author and providing 
'''					  this copyright info visible in your code and product 
'''					  documentation.
'''
''' DEPENDENCIES 	: None, except for being used together with the wrapper for the 
'''					  EA.Connector. The script should work 'as is' inside Enterprise 
'''					  Architect. 
''' TESTED			: Tested on Enterprise Architect 12.1 Beta, using the file
'''					  for simple property access
''' ===========================================================================

''' ---------------------------
''' CLASS CONNECTOR END WRAPPER
''' ---------------------------

Class TWConnectorEnd

	Dim m_eaConnectorEnd	As EA.ConnectorEnd
	Dim m_eaConnector		As EA.Connector
	Dim m_wconnector			''' As TWConnector
	Dim m_easubject			As EA.Element
	Dim m_SubjectID			''' As Long				''' Set on Wrap
	Dim m_eaElement			As EA.Element

	''' [TWConnectorEnd.Class_Initialize]
	Private Sub Class_Initialize()
		Set m_easubject = Nothing
		Set m_eaElement = Nothing
		Set m_wconnector = Nothing
		Set m_eaConnector = Nothing
		Set m_eaConnectorEnd = Nothing
	End Sub

	''' [TWConnectorEnd.Class_Terminate]
	Private Sub Class_Terminate()
		Set m_eaElement = Nothing
		Set m_eaConnectorEnd = Nothing
		Set m_eaConnector = Nothing
		Set m_wconnector = Nothing
		Set m_eaSubject = Nothing
	End Sub

	''' [TWConnectorEnd.Class_InitializeAfter]
	''' Called by the owning parent object, TWConnector in this case.
	Public Sub Class_InitializeAfter(aOwner)
		''' aOwner: EA.Connector
		''' Back-link to the Connector Wrapper. Useful during initialization.
		Set m_wconnector = aOwner
	End Sub


	''' [Aggregation]
	Public Property Get Aggregation()
		Aggregation = m_eaConnectorEnd.Aggregation
	End Property
	

	''' [Aggregation]
	''' Set this End's Aggregation AND the opposite End to it's logical kind!
	Public Property Let Aggregation(aAggregationKind)
		Dim OppositeEnd	''' TWConnectionEnd
		
		If IsAssociationKind Then
			If IsNearEnd then
				Set OppositeEnd = m_wconnector.m_otherEnd
			Else
				Set OppositeEnd = m_wconnector.m_nearEnd
			End If
			
			Select Case aAggregationKind
				Case EA_none
					''' Undoing aggregation doesn't affect the opposite end.
					m_eaConnectorEnd.Aggregation = EA_none
					
				Case EA_shared
					m_eaConnectorEnd.Aggregation = EA_shared
					''' Here the opposite end MUST be "reset"
					OppositeEnd.Aggregation = EA_none
				Case EA_composite
					m_eaConnectorEnd.Aggregation = EA_composite
					''' Here the opposite end MUST be "reset"
					OppositeEnd.Aggregation = EA_none
				Case Else
					Raise.Err err_AggregationKind, msg_AggregationKind
			End Select
			Update()
		End If
		
	End Property

	''' [IsAggregate]
	Public Property Get IsAggregate()
		if m_eaConnectorEnd.Aggregation = EA_shared then
			IsAggregate = True
		Else
			IsAggregate = False
		End If
	End Property

	''' [IsAggregationKind]
	Public Property Get IsAggregationKind()
		If m_eaConnectorEnd.Aggregation = EA_none Then
			IsAggregationKind = False		
		ElseIf m_eaConnectorEnd.Aggregation = EA_shared then
			IsAggregationKind = True
		ElseIf m_eaConnectorEnd.Aggregation = EA_composite Then
			IsAggregationKind = True
		Else
			Raise.Err err_AggregationKind, msg_AggregationKind
		End If
	End Property


	''' [IsComposite]
	Public Property Get IsComposite()
		if m_eaConnectorEnd.Aggregation = EA_composite then
			IsComposite = True
		Else
			IsComposite = False
		End If
	End Property
	

	''' [TWConnectorEnd.Wrap]
	''' aConnectorWrapper = TWConnector
	Public Function Wrap(aEAEnd, aEndClass) ''': TWConnector 
	
		Dim own_conn								''' Local vars speeds up
		
		''' Wrap/assign the EA objects
		Set own_conn = m_wconnector
		Set m_eaConnectorEnd = aEAEnd							''' The EA counterpart
		Set m_eaConnector = own_conn.m_eaConnector	''' The EA.Connector
		Set m_eaSubject = own_conn.m_eaSubject		''' The EA.Element currently using this wrapper
		m_SubjectID = own_conn.m_SubjectID			''' = m_eaSubject.ElementID 
		
		Set m_eaElement = aEndClass					''' The associated EA.Element object
				
		''' AutoName only if the AutoName property is set on the Connector.
		''' (because at this point in time we have no way of knowing the init 
		''' parameter for this. Therefore disabled here: /* EnsureRoleName() */		
		Set Wrap = Me
	End Function

	''' [End_]
	Public Property Get End_()
		End_ = m_eaConnectorEnd.End()		
	End Property

	''' [IsSingleLink]
	Public Property Get IsSingleLink()
		Dim Role As EA.ConnectorEnd
		
		Set Role = m_eaConnectorEnd
		IsSingleLink = (Role.Cardinality = "1")
		If Not IsSingleLink Then  ''' Check more.
			IsSingleLink = Role.Cardinality = "0..1"
		End If
	End Property

	''' [Cardinality]
	Public Property Get Cardinality()
		Cardinality = m_eaConnectorEnd.Cardinality
	End Property


	''' [eaClass]
	Public Property Get eaClass()
		Set eaClass = m_eaElement
	End Property


	''' [EndClassName]
	Public Property Get EndClassName()
		EndClassName = m_eaElement.Name()		
	End Property

	''' [IsAssociationKind]
	Public Property Get IsAssociationKind()
		IsAssociationKind = m_wconnector.IsAssociationKind
	End Property
	
	''' [Navigability]
	Public Property Get Navigability() ''': String
		Navigability = m_eaConnectorEnd.Navigable()					''' TODO: Scheck if working properly!
	End Property

	''' [IsNavigable]
	''' "Navigable" and "Unspecified" = "Navigable
	Public Property Get IsNavigable() ''': Boolean
		if m_eaConnectorEnd.Navigable = "Non-Navigable" then
			IsNavigable = False
		Else
			IsNavigable = True
		End If
	End Property

	''' [IsExplicitlyNavigable]
	''' Assuming that aslo Unspecified is "Navigable", then we need to also 
	''' check for explicit navigability
	Public Property Get IsExplicitlyNavigable() ''': Boolean
		if m_eaConnectorEnd.Navigable = "Navigable" then
		'if m_eaConnectorEnd.Navigability = "Navigable" then
			IsExplicitlyNavigable = True
		Else
			IsExplicitlyNavigable = False
		End If
	End Property


	''' [IsClientEnd]
	Public Property Get IsClientEnd() ''': Boolean
		IsClientEnd = End_ = "Client"
	End Property
	
	
	''' [IsSupplierEnd]
	Public Property Get IsSupplierEnd() ''': Boolean
		IsSupplierEnd = End_ = "Supplier"
	End Property


	''' [IsNearEnd]
	Public Property Get IsNearEnd()
		'' DEBUG
'		Session.Output	"m_SubjectID	    : " & m_SubjectID
'		Session.Output	"m_eaElement.ElementID: " & m_eaElement.ElementID
		
		IsNearEnd = m_SubjectID = m_eaElement.ElementID
	End Property
	
	''' [IsOtherEnd]
	Public Property Get IsOtherEnd()
		'' DEBUG
'		Session.Output	"m_SubjectID	    : " & m_SubjectID
'		Session.Output	"m_eaElement.ElementID: " & m_eaElement.ElementID
		
		IsOtherEnd = Not IsNearEnd
	End Property

	''' [Get TWConnectorEnd.Name]
	Public Property Get Name()
		Name = m_eaConnectorEnd.Role()			''' Role = Name
	End Property
	Public Property Let Name(aName)
		m_eaConnectorEnd.Role = aName			''' Role = Name
	End Property


	''' [Visibility]
	Public Property Get Visibility()
		Visibility = m_eaConnectorEnd.Visibility()
	End Property



	''' [TWConnectorEnd.Get eaConnector]
	Public Property Get eaConnector()
		Set eaConnector = m_eaConnector
	End Property	


	''' [TWConnectorEnd.Get eaConnectorEnd]
	Public Property Get eaConnectorEnd()
		Set eaConnectorEnd = m_eaConnectorEnd
	End Property
	Private Property Let eaConnectorEnd(aRole)
		Set m_eaConnectorEnd = aRole
	End Property


	''' [TWConnectorEnd.EnsureRoleName]
	''' Applicable only to IsAssociationKinds. Ensure also that
	''' names are created only from NearEnd = CliendEnd (owning side)
	''' as to result in proper order of names in the concatenated Link name
	Public Sub EnsureRoleName()
		''' Nearest End - AutoName if the role name is missing.
		
		if Len(Name) = 0 then
			''' Make new name based on the CLASSNAME of the referenced class
			sRoleName = EndClassName()
			Name = LCase(Mid(sRoleName, 1, 1)) & Mid(sRoleName, 2, Len(sRoleName)-1) ''' DelphiCase->camelCase
			Update()
			Refresh()
			Session.Output "WARNING: Auto-named role name : " & sRoleName			'''(($test))
		end if
	End Sub
	
	
	''' [Update]
	Public Sub Update()
		m_eaConnectorEnd.Update()
	End Sub

	''' [Refresh]
	Public Sub Refresh()
		''' m_wconnector.DiagramID
	End Sub



	''' [TWConnectorEnd.]
	
	
	'' ROLE END
End Clas