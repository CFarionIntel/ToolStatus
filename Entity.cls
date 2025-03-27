VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Entity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Entity Class Module Member Variables
Private m_STRING_Name As String
Private m_STRING_Availability As String
Private m_STRING_State As String
Private m_STRING_CEID As String
Private m_STRING_VFMFGID As String

'Entity Class Module Properties
Public Property Get Name() As String
    Name = m_STRING_Name
End Property

Public Property Let Name(Value As String)
    m_STRING_Name = Value
End Property

Public Property Get Availability() As String
    Availability = m_STRING_Availability
End Property

Public Property Let Availability(Value As String)
    m_STRING_Availability = Value
End Property

Public Property Get State() As String
    State = m_STRING_State
End Property

Public Property Let State(Value As String)
    m_STRING_State = Value
End Property

Public Property Get CEID() As String
    CEID = m_STRING_CEID
End Property

Public Property Let CEID(Value As String)
    m_STRING_CEID = Value
End Property

Public Property Get Tools() As String
    Tools = Etch_Tools
End Property

Public Property Let Tools(Value As String)
    Etch_Tools = Value
End Property

Public Property Get VFMFGID() As String
    VFMFGID = m_STRING_VFMFGID
End Property

Public Property Let VFMFGID(Value As String)
    m_STRING_VFMFGID = Value
End Property

'Entity Class Module Events
Private Sub Class_Initialize()
    m_STRING_Name = ""
    m_STRING_Availability = ""
    m_STRING_State = ""
    m_STRING_CEID = ""
End Sub

'Entity Class Module Methods
Public Function Verify_Entity() As Boolean
    Dim ret As Boolean
    
    Dim Chars_in_MOM As Long
    Dim Chars_in_Child As Long
    Dim Left_Child_Char As String
    Dim Cu_Chamber As String
    Dim Chamber_Number As Long
    Dim Tool_Codes As String
    
    Dim Legit_Tool As Boolean
    
    ret = False
    
    Components = Split(m_STRING_Name, "_")
    'Make sure it is not blank or has more than child components
    If UBound(Components) < 0 Or UBound(Components) > 1 Then
        Verify_Entity = ret
        Exit Function
    End If
    
    'Check if MOM is in the correct format
    Chars_in_MOM = Len(Components(0))
    If Chars_in_MOM < 5 Or Chars_in_MOM > 6 Then
        Verify_Entity = ret
        Exit Function
    End If
    
    '3 letter abbreviations for tools here?
    Tool_Codes = "AER;AME;ANT;APT;BCL;BET;CAR;D1N;D1S;DCR;DEA;EVA;FAW;GBB;GIB;GNT;GNX;GPC;GRX;GSX;GTA;GTO;GTR;GTT;GTX;HIN;HME;HOP;ICE;LAT;MHU;MIK;MME;MPT;NAB;ONT;OVM;OVN;OVO;OVP;OVR;OVS;OXS;OXT;PAT;PED;PGX;RCL;REX;SLS;TAO;TEX;TLR;TLS;TVO;UFP"
    
    Legit_Tool = Tool_Abbreviation_Check(Tool_Codes, Left(Components(0), 3))
    If Legit_Tool <> True Then
        Verify_Entity = ret
        Exit Function
    End If
    
    If UBound(Components) = 1 Then
        'Check if Child is in the correct format
        Chars_in_Child = Len(Components(1))
        If Chars_in_Child <> 3 Then
            Verify_Entity = ret
            Exit Function
        End If
        
        Left_Child_Char = Left(Components(1), 1)
        If Left_Child_Char <> "P" Then
            Verify_Entity = ret
            Exit Function
        End If
        
        Cu_Chamber = Mid(Components(1), 2, 1)
        If Cu_Chamber <> "C" And Cu_Chamber <> "M" Then
            Verify_Entity = ret
            Exit Function
        End If
        
        Chamber_Number = Int(Right(Components(1), 1))
        If Chamber_Number < 1 Or Chamber_Number > 8 Then
            Verify_Entity = ret
            Exit Function
        End If
    End If
    ret = True
    Verify_Entity = ret
End Function

Private Function Tool_Abbreviation_Check(Tool_Codes As String, Tool_in_Question As String) As Boolean
    Dim ret As Boolean
    Dim lower_bound As Long
    Dim upper_bound As Long
    
    
    ret = False
    'ret = True ' Temp while adding feature to code
    
    Tool_Codes_Array = Split(Tool_Codes, ";")
    
    Num_of_Tool_Codes = UBound(Tool_Codes_Array)
    
    Dim Current_Row As Long
    Dim current_value As String
    
    lower_bound = 0
    upper_bound = Num_of_Tool_Codes
    
    While lower_bound < upper_bound And (lower_bound + 1) <> upper_bound
        Current_Row = (lower_bound + upper_bound) / 2
        current_value = Tool_Codes_Array(Current_Row)
        If current_value = Tool_in_Question Then
            'Exit Condition for matching names
            lower_bound = upper_bound + 1
        ElseIf current_value < Tool_in_Question Then
            lower_bound = Current_Row
        ElseIf current_value > Tool_in_Question Then
            upper_bound = Current_Row
        End If
    Wend
    
    If (current_value = Tool_in_Question) Then
        ret = True
    Else
        ret = False
    End If
    
    Tool_Abbreviation_Check = ret
End Function
