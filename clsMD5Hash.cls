VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMD5Hash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'
' MD5Hash
'
' Perform CryptoAPI MD5 hash of contents of a named file or a Byte array,
' returning hash as String of 32 hex digits.

' ----- Private Consts -----

Private Const ALG_TYPE_ANY          As Long = 0
Private Const ALG_CLASS_HASH        As Long = 32768
Private Const ALG_SID_MD5           As Long = 3
Private Const CALG_MD5              As Long = ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD5

Private Const PROV_RSA_FULL         As Long = 1
Private Const CRYPT_VERIFYCONTEXT   As Long = &HF0000000
Private Const MS_DEFAULT_PROVIDER   As String = _
    "Microsoft Base Cryptographic Provider v1.0"

Private Const HP_HASHVAL            As Long = 2
Private Const HP_HASHSIZE           As Long = 4

' ----- Private Defines -----

Private Declare Function CryptAcquireContext Lib "advapi32" Alias "CryptAcquireContextA" ( _
    ByRef phProv As Long, _
    ByVal pszContainer As String, _
    ByVal pszProvider As String, _
    ByVal dwProvType As Long, _
    ByVal dwFlags As Long) As Long ' TRUE (<> 0) = success.  See Err.LastDLLError if FALSE.

Private Declare Function CryptCreateHash Lib "advapi32" ( _
    ByVal hProv As Long, _
    ByVal algid As Long, _
    ByVal hKey As Long, _
    ByVal dwFlags As Long, _
    ByRef phHash As Long) As Long ' TRUE (<> 0) = success.  See Err.LastDLLError if FALSE.
    
Private Declare Function CryptDestroyHash Lib "advapi32" ( _
    ByVal hHash As Long) As Long ' TRUE (<> 0) = success.  See Err.LastDLLError if FALSE.

Private Declare Function CryptGetHashParam Lib "advapi32" ( _
    ByVal hHash As Long, _
    ByVal dwParam As Long, _
    ByRef pbData As Any, _
    ByRef pdwDataLen As Long, _
    ByVal dwFlags As Long) As Long

Private Declare Function CryptHashData Lib "advapi32" ( _
    ByVal hHash As Long, _
    ByRef pbData As Any, _
    ByVal dwDataLen As Long, _
    ByVal dwFlags As Long) As Long

Private Declare Function CryptReleaseContext Lib "advapi32" ( _
    ByVal hProv As Long, _
    ByVal dwFlags As Long) As Long ' TRUE (<> 0) = success.  See Err.LastDLLError if FALSE.

'----- Private Data -----

Private m_hHash As Long     ' Hash object handle.
Private m_hProvider As Long ' Cryptographic Service Provider handle.

'----- Private Methods -----

Private Sub HashBlock(ByRef Block() As Byte)
    If CryptHashData(m_hHash, _
                     Block(LBound(Block)), _
                     UBound(Block) - LBound(Block) + 1, _
                     0&) = 0 Then
        Err.Raise vbObjectError Or &HC312&, _
                  "clsMD5Hash", _
                  "Failed to hash data block, system error " _
                & CStr(Err.LastDllError)
    End If
End Sub

Private Function HashValue() As Byte()
    Dim lngDataLen As Long
    Dim lngHashSize As Long
    Dim bytHashValue() As Byte
    
    lngDataLen = 4 ' 4 bytes for Long length.
    
    If CryptGetHashParam(m_hHash, HP_HASHSIZE, lngHashSize, lngDataLen, 0&) = 0 Then
        Err.Raise vbObjectError Or &HC322&, _
                  "clsMD5Hash", _
                  "Failed to obtain hash value length, system error " _
                & CStr(Err.LastDllError)
    Else
        lngDataLen = lngHashSize
        ReDim bytHashValue(lngDataLen - 1)
        
        If CryptGetHashParam(m_hHash, HP_HASHVAL, bytHashValue(0), lngDataLen, 0&) = 0 Then
            Err.Raise vbObjectError Or &HC324&, _
                      "clsMD5Hash", _
                      "Failed to obtain hash value, system error " _
                    & CStr(Err.LastDllError)
        Else
            HashValue = bytHashValue
            
            CryptDestroyHash m_hHash
        End If
    End If
End Function

Private Sub NewHash()
Attribute NewHash.VB_Description = "Initialize a new hashing sequence"
    If CryptCreateHash(m_hProvider, CALG_MD5, 0&, 0&, m_hHash) = 0 Then
        Err.Raise vbObjectError Or &HC332&, _
                  "clsMD5Hash", _
                  "Failed to create CryptoAPI Hash object, system error " _
                & CStr(Err.LastDllError)
    End If
End Sub

'----- Public Methods -----

Public Function HashBytes(ByRef Block() As Byte) As Byte()
    NewHash
    HashBlock Block
    HashBytes = HashValue()
End Function

'----- Class Event Handlers -----

Private Sub Class_Initialize()
    If CryptAcquireContext(m_hProvider, _
                           vbNullString, _
                           MS_DEFAULT_PROVIDER, _
                           PROV_RSA_FULL, _
                           CRYPT_VERIFYCONTEXT) = 0 Then
        Err.Raise vbObjectError Or &HC352&, _
                  "clsMD5Hash.Class_Initialize", _
                  "Failed to obtain access to CryptoAPI, system error " _
                & CStr(Err.LastDllError)
    End If
End Sub

Private Sub Class_Terminate()
    On Error Resume Next    ' All exceptions must be processed here.
    
    CryptDestroyHash m_hHash
    CryptReleaseContext m_hProvider, 0&
End Sub
