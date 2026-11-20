Imports System
Imports System.Runtime.InteropServices
Imports System.Security.Cryptography.X509Certificates

Public Module SignatureCheck

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
    Private Structure WINTRUST_FILE_INFO
        Public cbStruct As UInteger
        Public pcwszFilePath As String
        Public hFile As IntPtr
        Public pgKnownSubject As IntPtr
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Private Structure WINTRUST_DATA
        Public cbStruct As UInteger
        Public pPolicyCallbackData As IntPtr
        Public pSIPClientData As IntPtr
        Public dwUIChoice As UInteger
        Public fdwRevocationChecks As UInteger
        Public dwUnionChoice As UInteger
        Public pFile As IntPtr
        Public dwStateAction As UInteger
        Public hWVTStateData As IntPtr
        Public pwszURLReference As IntPtr
        Public dwProvFlags As UInteger
        Public dwUIContext As UInteger
        Public pSignatureSettings As IntPtr
    End Structure

    <DllImport("wintrust.dll", ExactSpelling:=True, CharSet:=CharSet.Unicode)>
    Private Function WinVerifyTrust(hwnd As IntPtr, <MarshalAs(UnmanagedType.LPStruct)> pgActionID As Guid, ByRef pWVTData As WINTRUST_DATA) As Integer
    End Function

    ''' <summary>
    ''' True only if the file has a valid Authenticode signature (hash intact, chain trusted by Windows)
    ''' and the signer is Microsoft Corporation.
    ''' </summary>
    Public Function IsSignedByMicrosoft(filePath As String) As Boolean
        Dim fileInfo As New WINTRUST_FILE_INFO With {
            .cbStruct = CUInt(Marshal.SizeOf(GetType(WINTRUST_FILE_INFO))),
            .pcwszFilePath = filePath
        }
        Dim pFile As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(GetType(WINTRUST_FILE_INFO)))
        Try
            Marshal.StructureToPtr(fileInfo, pFile, False)
            ' dwUIChoice 2 = WTD_UI_NONE, dwUnionChoice 1 = WTD_CHOICE_FILE
            Dim data As New WINTRUST_DATA With {
                .cbStruct = CUInt(Marshal.SizeOf(GetType(WINTRUST_DATA))),
                .dwUIChoice = 2UI,
                .dwUnionChoice = 1UI,
                .pFile = pFile
            }
            ' WINTRUST_ACTION_GENERIC_VERIFY_V2
            Dim action As New Guid("00AAC56B-CD44-11d0-8CC2-00C04FC295EE")
            If WinVerifyTrust(New IntPtr(-1), action, data) <> 0 Then Return False
        Finally
            Marshal.DestroyStructure(pFile, GetType(WINTRUST_FILE_INFO))
            Marshal.FreeHGlobal(pFile)
        End Try

        ' Signature is valid and trusted, now make sure it is Microsoft's
        Try
            Dim cert As New X509Certificate2(X509Certificate.CreateFromSignedFile(filePath))
            Return cert.Subject.Contains("O=Microsoft Corporation")
        Catch ex As Exception
            Return False
        End Try
    End Function

End Module
