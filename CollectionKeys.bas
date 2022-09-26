Attribute VB_Name = "CollectionKeys"
Option Explicit

#If Win64 Then
Private Declare PtrSafe Sub MemCopy Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As LongPtr, ByVal Source As LongPtr, ByVal Length As LongPtr)
#Else
Private Declare Sub MemCopy Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
#End If

#If Win64 Then
Public Function GetKeys(ByVal oColl As Collection) As String()
    Dim CollPtr As LongPtr
    Dim KeyPtr As LongPtr
    Dim ItemPtr As LongPtr

    'Get MemoryAddress of Collection Object
    CollPtr = ObjPtr(oColl)

    'Peek ElementCount
    Dim ElementCount As Long
    ElementCount = PeekLong(CollPtr + 28)

    'Verify ElementCount
    If ElementCount <> oColl.Count Then
        Stop
    End If

    Dim index As Long
    Dim Temp() As String
    ReDim Temp(ElementCount) As String

    'Get MemoryAddress of first CollectionItem
    ItemPtr = PeekLongLong(CollPtr + 40)

    'Loop through all CollectionItems in Chain
    While Not ItemPtr = 0 And index < ElementCount
        index = index + 1

        'Get MemoryAddress of Element-Key
        KeyPtr = PeekLongLong(ItemPtr + 24)

        'Peek Key and add to temporary array (if present)
        If KeyPtr <> 0 Then
           Temp(index) = PeekBSTR(KeyPtr)
        End If

        'Get MemoryAddress of next Element in Chain
        ItemPtr = PeekLongLong(ItemPtr + 40)
    Wend

    GetKeys = Temp
End Function

Private Function PeekLong(Address As LongPtr) As Long
    If Address = 0 Then Stop
  
    MemCopy VarPtr(PeekLong), Address, 4^
End Function

Private Function PeekLongLong(Address As LongPtr) As LongLong
    If Address = 0 Then Stop
  
    MemCopy VarPtr(PeekLongLong), Address, 8^
End Function

Private Function PeekBSTR(Address As LongPtr) As String
    Dim Length As Long

    If Address = 0 Then Stop
    
    Length = PeekLong(Address - 4)
    PeekBSTR = Space$(Length \ 2)
    
    MemCopy StrPtr(PeekBSTR), Address, CLngLng(Length)
End Function

#Else

Public Function GetKeys(ByVal oColl As Collection) As String()
    Dim CollPtr As Long
    Dim KeyPtr As Long
    Dim ItemPtr As Long

    'Get MemoryAddress of Collection Object
    CollPtr = ObjPtr(oColl)

    'Peek ElementCount
    Dim ElementCount As Long
    ElementCount = PeekLong(CollPtr + 16)

    If ElementCount <> oColl.Count Then
        Stop
    End If

    Dim index As Long
    Dim Temp() As String
    ReDim Temp(ElementCount)

    'Get MemoryAddress of first CollectionItem
    ItemPtr = PeekLong(CollPtr + 24)

    While Not ItemPtr = 0 And index < ElementCount
        index = index + 1

        'Get MemoryAddress of Element-Key
        KeyPtr = PeekLong(ItemPtr + 16)

        'Peek Key and add to temporary array (if present)
        If KeyPtr <> 0 Then
           Temp(index) = PeekBSTR(KeyPtr)
        End If

        'Get MemoryAddress of next Element in Chain
        ItemPtr = PeekLong(ItemPtr + 24)
    Wend

    GetKeys = Temp
End Function

Private Function PeekLong(Address As Long) As Long
  If Address = 0 Then Stop
  
  MemCopy VarPtr(PeekLong), Address, 4&
End Function

Private Function PeekBSTR(Address As Long) As String
    Dim Length As Long

    If Address = 0 Then Stop
    
    Length = PeekLong(Address - 4)
    PeekBSTR = Space(Length \ 2)
    
    MemCopy StrPtr(PeekBSTR), Address, Length
End Function

#End If
