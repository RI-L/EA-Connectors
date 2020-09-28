'Option Explicit

!INC Local Scripts.EAConstants-VBScript
!INC Connectors.ConnectorEnd
!INC TaggedValues.Wrappers_Const


''' ===========================================================================
''' CONNECTOR WRAPPER
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



Dim m_conapi 		''' You may want to use this variable directly after initialized
					''' in the first access of the ConApi pseudo 'property'.

Public Function ConApi()
	''' Ensure that the TaggedVaslue helper is created only once.
	if m_conapi is Nothing then
		Set m_conapi = New TWConnector
	End If
	Set ConApi = m_conapi
End Function

Private Sub Initialize_ConApi()
	''' Ensure initialization of the variable as to prepare 
	''' for the assigment check in the ConApi() function
	Set m_conapi = Nothing
End Sub

Private Sub Module_Terminate()
	''' Call this explicitly to dispose of the object.
	Set m_conapi = Nothing
End Sub


''' ---------------------------------------------------------------------------


Class TWConnector

	Dim m_eaSubject 	As EA.Element			''' The object currently using the wrapper
	Dim m_SubjectID								''' = m_eaSubject.ElementID
	Dim m_SubjectGUID							''' = m_eaSubject.ElementGUID
	Dim m_eaConnector			As EA.Connector			''' The wrapped EA.Connector counter part
	
	Dim m_nearEnd		''' As TWConnectorEnd
	Dim m_otherEnd		''' As TWConnectorEnd
	Dim m_owningEnd		''' As TWConnectorEnd 	''' = ClientEnd / Source
	
	Dim m_IsAutoName	''' As Boolean
	
	
	''' [TWConnector.Class_Initialize]
	Private Sub Class_Initialize()
		Set m_eaSubject = Nothing
		Set m_eaConnector = Nothing
		Set m_owningEnd = Nothing
		m_IsAutoName = False
		
		Set m_nearEnd = New TWConnectorEnd
		m_nearEnd.Class_InitializeAfter(Me)
		
		Set m_otherEnd  = New TWConnectorEnd
		m_otherEnd.Class_InitializeAfter(Me)
	End Sub


	''' [TWConnector.Class_Terminate]
	Private Sub Class_Terminate()
		Set m_eaConnector = Nothing
		Set m_eaSubject = Nothing
		Set m_otherEnd = Nothing
		Set m_nearEnd = Nothing
		Set m_owningEnd = Nothing
	End Sub

	
	''' [TWConnector.Class_InitializeAfter]
	''' Called by the owning parent object, TWConnector in this case.
	Public Sub Class_InitializeAfter(aOwner)
		''' aOwner: EA.Connector
		Set m_owner = aOwner
	End Sub


	''' [TWConnector.Wrap]
	''' Param: aSubjectElement: EA.Element
	''' Param: aConnector: EA.Connector
	Public Function Wrap(ByRef aSubjectElement, ByRef aEaConnector) ''': TWConnector
		
		Set m_eaSubject = aSubjectElement
		
		''' Store Subject's ElementID
		Select Case aSubjectElement.ObjectType
			Case EA_Class, EA_Interface
				m_SubjectID 	= aSubjectElement.ElementID
				m_SubjectGUID 	= aSubjectElement.ElementGUID
			Case EA_Package
				Dim elem
				Set elem 		= aSubjectElement.Element
				m_SubjectID 	= elem.ElementID 
				m_SubjectGUID 	= elem.ElementGUID
			Case Else
					m_SubjectID = 0
					m_SubjectGUID = ""
					Stop
		End Select
		
		Set m_eaConnector = aEaConnector
		
		''' WrapConnectionEnds must always be called last, since the end roles
		''' uses some of the WConnectors settings to retrieve more info
		WrapConnectionEnds()
		
		''' Return Self
		Set Wrap = Me
	End Function


	''' [WrapConnectionEnds]
	''' Connect the nearest EndRole to "nearRole", and the "owning" role (the source)
	''' to owningEnd
	Private Sub WrapConnectionEnds()
		''' Local vars speeds things up
		Dim OtherObj As EA.Element
		Dim c As EA.Connector
		Dim NearE As EA.ConnectorEnd
		Dim OtherE As EA.ConnectorEnd
		
		Set c = m_eaConnector
		''' If Subject is Client/Source, then make that "Near End"
		If m_SubjectID = c.ClientID then
			''' NEAR END = CLIENT
			Set NearE = c.ClientEnd
			Set OtherE = c.SupplierEnd
			''' Supplier Objet in the other end
			Set OtherObj = Repository.GetElementByID(c.SupplierID)
		ElseIf m_SubjectID = c.SupplierID Then
			''' NEAR END = SUPPLIER
			Set NearE = c.SupplierEnd
			Set OtherE = c.ClientEnd
			''' Client Object in the other end
			Set OtherObj = Repository.GetElementByID(c.ClientID)
		Else
			Raise.Err err_UnknownSubjectID, msg_UnknownSubjectID
		End If
		''' WRAP - Subject must always be NearEnd
		m_nearEnd.Wrap NearE, m_eaSubject			
		m_otherEnd.Wrap OtherE, OtherObj

		''' OWNING END - In any case, the Client side must always (technically 
		''' speaking) be on the "owning" side.
		Set m_owningEnd = c.ClientEnd

		''' All role ends should now be wrapped.
	End Sub

	''' [SetNearEndAggregation]
	Public Function SetNearEndAggregation(aAggregationKind)
		m_nearEnd.Aggregation = aAggregationKind
		Set SetNearEndAggregation = Me			''' Sets a param om create...
	End Function

	''' [SetOtherEndAggregation]
	Public Function SetOtherEndAggregation(aAggregationKind)
		m_otherEnd.Aggregation = aAggregationKind
		Set SetOtherEndAggregation = Me			''' Sets a param om create...
	End Function


	''' END Initialize methods
	
	''' [IsOwnedByNearEnd]
	Public Property Get IsOwnedByNearEnd()
		IsOwnedByNearEnd = m_SubjectGUID = m_nearEnd.m_eaElement.ElementGUID
	End Property
	
	
	''' [Direction]
	Public Property Get Direction()
		Direction = m_eaConnector.Direction
	End Property
	
	''' [IsAggregationKind]
	''' True if either end is either Aggregate or Composite
	Public Property Get IsAggregationKind()
		if m_nearEnd.IsAggregationKind then
			IsAggregationKind = True
		ElseIF m_otherEnd.IsAggregationKind Then
			IsAggregationKind = True
		Else
			IsAggregationKind = False
		End If
	End Property


	''' [IsAggregate]
	''' Checks both ends if any of them is an Aggregate
	Public Property Get IsAggregate()
		if m_nearEnd.IsAggregate then
			IsAggregate = True
		ElseIF m_otherEnd.IsAggregate Then
			IsAggregate = True
		Else
			IsAggregate = False
		End If
	End Property


	''' [IsComposite]
	''' Checks both ends if any of them is a Composite
	Public Property Get IsComposite()
		if m_nearEnd.IsComposite then
			IsComposite = True
		ElseIF m_otherEnd.IsComposite Then
			IsComposite = True
		Else
			IsComposite = False
		End If
	End Property


	''' [EnsureConnectionName]
	''' Auto create connector name (based on roles). If the link lacks a name 
	'''	this method concatenates the two role names and clips the name to maximum 
	''' length 32 chars.
	''' Attention: EnsureRoleNames() MUST be called BEFORE this method, since
	''' it uses the Role Names to create the Link name.
	Sub EnsureConnectionName()
		''' Only if name is empty
		if Len(m_eaConnector.Name()) = 0 Then
			''' Don't mess with Generalizations and Dependencies and stuff
			''' Applies only to IsAssociationKinds!
			Dim sNear
			Dim sOther
			
			''' Create the name in this order. Optimally only could use 
			''' Source/ClientEnd as the first part. We'll see that
			sNear = nearEnd.Name()
			sOther = otherEnd.Name()
			
			If Len(sNear)>0 then
			''' Remove up and past underscores, like "xx_|name" :
			''' Make the string "Captialized" & "Captialized"
				If InStr(1, sNear, "_", 1) > 1 Then _
					sNear = Mid(sNear, InStr(1, sNear, "_", 1)+1, 1000)
				sNear = UCase(Mid(sNear, 1, 1)) & Mid(sNear, 2, Len(sNear)-1) ''' DelphiCase->camelCase
			End If
			
			If Len(sOther)>0 then
				If InStr(1, sOther, "_", 1) > 1 Then _
					sOther = Mid(sOther, InStr(1, sOther, "_", 1)+1, 1000)
				sOther = UCase(Mid(sOther, 1, 1)) & Mid(sOther, 2, Len(sOther)-1) ''' DelphiCase->camelCase
			End If
			''' Insert new name clipped at max 32 char long str
			If Len(sNear & sOther)>0 then
				S = Left("Lnk" & sNear & "_" & sOther, 32)
				m_eaConnector.Name = S
				
				''' Update model
				m_eaConnector.Update()
				Session.Output "WARNING: Auto-named Link: " & Name 				''' (($test))
			Else
				S = "ERROR: Attempt to Auto-name Link failed: ("				''' (($test))
				S = S & m_eaConnector.Type & ", for class: " & eaSubject.Name() ''' (($test))
				Session.Output S
			End If
		End if
	End Sub


	''' [EnsureRoleNames]
	Sub EnsureRoleNames() 
		''' Perform AutoNaming, but not for generalizations and stuff
		if m_IsAutoName then
			if IsAssociationKind then
				''' only if empty
				If Len(m_nearEnd.Name()) = 0 Then _
					m_nearEnd.EnsureRoleName()
				If Len(m_otherEnd.Name()) = 0 Then _
					m_otherEnd.EnsureRoleName()
			End If
		End If
	End Sub


	''' [AutoName]
	''' Sets IsAutoName (may be needed if any role or connection name is empty). 
	''' Called like so: sName = Wrap(cls, conn).AutoName(true/false).Name()
	Public Function AutoName(aDoAutoName)
		m_IsAutoName = aDoAutoName
		''' Rename only if "owning" the link (NearEnd = ClientEnd).
		''' (Stepwise IFs below due to performance reasons)
		If aDoAutoName Then
			If IsAssociationKind then
				If IsOwnedByNearEnd Then
					''' Important to AutoName the roles first, because the Connection Name
					''' is composed of the two role names combined.
					EnsureRoleNames()
					EnsureConnectionName()
				End If
			End If
		End If
		''' Allow chaining this command
		Set AutoName = Me			''' Sets a param om create...
	End Function
	
	''' [IsAutoName]
	Public Property Get IsAutoName() ''': Boolean
		IsAutoName = m_IsAutoName
	End Property
	

	''' [IsNearEndAggregate]
	Public Property Get IsNearEndAggregate() ''': Boolean
		IsNearEndAggregate = m_AggregatonKind
	End Property

	
	''' PROPERTIES
	''' PUBLIC


	''' [Name]
	''' Internal pointer to EA.Connector should have been assigned in Wrap()
	Public Property Get Name() ''': String
		Name = m_eaConnector.Name()
	End Property


	''' [Get eaSubject]
	Public Property Get eaSubject()
		Set eaSubject = m_eaSubject
	End Property


	''' [IsAssociation]
	Public Property Get IsAssociation() ''': Boolean
		IsAssociation = m_eaConnector.Type() = METATYPE_ASSOCIATION
	End Property
	
	
	''' [IsAssociationKind]
	Public Property Get IsAssociationKind() ''': Boolean
		Select Case m_eaConnector.Type
			Case METATYPE_ASSOCIATION, METATYPE_AGGREGATION, METATYPE_COMPOSITION
				IsAssociationKind = True
			Case Else
				IsAssociationKind = False
			End Select
	End Property

	''' [IsGeneralization]
	Public Property Get IsGeneralization() ''': Boolean
		IsGeneralization = m_eaConnector.Type = METATYPE_GENERALIZATION
	End Property

	''' [IsRealisation]
	Public Property Get IsRealisation() ''': Boolean
		IsRealisation = m_eaConnector.Type = METATYPE_REALISATION
	End Property

	''' [IsDependency]
	Public Property Get IsDependency() ''': Boolean
		IsDependency = m_eaConnector.Type = METATYPE_DEPENDENCY
	End Property
	
	''' [IsAggregation]
	Public Property Get IsAggregation() ''': Boolean
		IsAggregation = m_eaConnector.Type = METATYPE_AGGREGATION
	End Property

	''' [IsComposition]
	Public Property Get IsComposition() ''': Boolean
		IsComposition = m_eaConnector.Type = METATYPE_COMPOSITION
	End Property
	

	''' [nearEnd]
	Public Property Get nearEnd()
		Set nearEnd = m_nearEnd
	End Property


	''' [otherEnd]
	Public Property Get otherEnd()
		Set otherEnd = m_otherEnd
	End Property


	''' [Type_]
	''' The EA property .type is a keyword in VBScript, so we add an
	''' underscore to mask the invalid name
	Public Property Get Type_() ''': String
		Type_ = m_eaConnector.Type()
	End Property


	''' [Get eaConnector]
	Public Property Get eaConnector()
		Set eaConnector = m_eaConnector
		
	End Property

End Class ''' CONNECTOR WRAPPER


Initialize_ConAp