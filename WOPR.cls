VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WOPR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Entity Class Module Member Variables
Private LONG_ID As Long
Private STRING_WOPR_Status As String
Private STRING_ENTITY As String
Private STRING_CEID As String
Private STRING_ENTITY_STATE As String
Private LONG_Prio As Long
Private STRING_Created_Date As String
Private STRING_Last_Updated As String
Private STRING_Description As String

Private Sub Class_Initialize()
    LONG_ID = 0
    STRING_WOPR_Status = ""
    STRING_ENTITY_STATE = ""
    STRING_ENTITY = ""
    STRING_CEID = ""
    LONG_Prio = 5
    STRING_Created_Date = ""
    STRING_Last_Updated = ""
    STRING_Description = ""
End Sub

'Entity Class Module Properties
Public Property Get ID() As Long
    ID = LONG_ID
End Property

Public Property Let ID(Value As Long)
    LONG_ID = Value
End Property

Public Property Get Sts() As String
    Sts = STRING_WOPR_Status
End Property

Public Property Let Sts(Value As String)
    STRING_WOPR_Status = Value
End Property

Public Property Get State() As String
    State = STRING_ENTITY_STATE
End Property

Public Property Let State(Value As String)
    STRING_ENTITY_STATE = Value
End Property

Public Property Get Entity() As String
    Entity = STRING_ENTITY
End Property

Public Property Let Entity(Value As String)
    STRING_ENTITY = Value
End Property

Public Property Get CEID() As String
    CEID = STRING_CEID
End Property

Public Property Let CEID(Value As String)
    STRING_CEID = Value
End Property

Public Property Get Prio() As Long
    Prio = LONG_Prio
End Property

Public Property Let Prio(Value As Long)
    LONG_Prio = Value
End Property

Public Property Get CreatedDate() As String
    CreatedDate = STRING_Created_Date
End Property

Public Property Let CreatedDate(Value As String)
    STRING_Created_Date = Value
End Property

Public Property Get LastUpdated() As String
    LastUpdated = STRING_Last_Updated
End Property

Public Property Let LastUpdated(Value As String)
    STRING_Last_Updated = Value
End Property

Public Property Get Description() As String
    Description = STRING_Description
End Property

Public Property Let Description(Value As String)
    STRING_Description = Value
End Property

Public Function Verify_WOPR() As Boolean
    Dim ret As Boolean
    
    Dim SubFab_WOPR As Long
    
    Dim Chars_in_MOM As Long
    Dim Tool_Exists As Boolean
    
    Dim Chars_in_Child As Long
    Dim Left_Child_Char As String
    Dim Cu_Chamber As String
    Dim Chamber_Number As Long
    Dim Etch_Chamber_String As String
    
    'This section I decided not to pursue. If the tool does not exist, it will get flushed in a later section
    Etch_Chamber_String = "AER;AME;ANT;APT;BCL;BET;CAR;D1N;D1S;DCR;DEA;EVA;FAW;GBB;GIB;GNT;GNX;GPC;GRX;GSX;GTA;GTO;GTR;GTT;GTX;HIN;HME;HOP;ICE;LAT;MHU;MIK;MME;MPT;NAB;ONT;OVM;OVN;OVO;OVP;OVR;OVS;OXS;OXT;PAT;PED;PGX;RCL;REX;SLS;TAO;TEX;TLR;TLS;TVO;UFP"
    Etch_Chambers = Split(Etch_Chamber_String, ";")
    Tool_Exists = False
    'Section Above
        
    ret = False
    SubFab_WOPR = 0
    
    Components = Split(STRING_ENTITY, "_")
    'Make sure it is not blank or has more than child components
    If Components(0) = "SF" Then
        SubFab_WOPR = 1
    End If
    
    If UBound(Components) < 0 Or UBound(Components) > (1 + SubFab_WOPR) Then
        Verify_Entity = ret
        Exit Function
    End If
    
    'Check if MOM is in the correct format
    Chars_in_MOM = Len(Components(0 + SubFab_WOPR))
    If Chars_in_MOM < 5 Or Chars_in_MOM > 6 Then
        Verify_Entity = ret
        Exit Function
    End If
    
    If UBound(Components) = (1 + SubFab_WOPR) Then
        'Check if Child is in the correct format
        Chars_in_Child = Len(Components(1 + SubFab_WOPR))
        If Chars_in_Child <> 3 Then
            Verify_Entity = ret
            Exit Function
        End If
        
        On Error Resume Next 'In case the second component of the work order is "LLM", "LP", or "VTM"
            Chamber_Number = Int(Right(Components(1 + SubFab_WOPR), 1))
        If Chamber_Number < 1 Or Chamber_Number > 8 Then
            Verify_Entity = ret
            Exit Function
        End If
    End If
        
    ret = True
    Verify_WOPR = ret
End Function

