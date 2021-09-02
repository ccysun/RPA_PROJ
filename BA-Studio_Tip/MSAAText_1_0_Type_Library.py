# -*- coding: mbcs -*-
# Created by makepy.py version 0.5.01
# By python version 3.9.1 (tags/v3.9.1:1e5d33e, Dec  7 2020, 17:08:21) [MSC v.1927 64 bit (AMD64)]
# From type library 'msaatext.dll'
# On Mon Jun  7 15:37:28 2021
'MSAAText 1.0 Type Library'
makepy_version = '0.5.01'
python_version = 0x30901f0

import win32com.client.CLSIDToClass, pythoncom, pywintypes
import win32com.client.util
from pywintypes import IID
from win32com.client import Dispatch

# The following 3 lines may need tweaking for the particular server
# Candidates are pythoncom.Missing, .Empty and .ArgNotFound
defaultNamedOptArg=pythoncom.Empty
defaultNamedNotOptArg=pythoncom.Empty
defaultUnnamedArg=pythoncom.Empty

CLSID = IID('{150E2D7A-DAC1-4582-947D-2A8FD78B82CD}')
MajorVersion = 1
MinorVersion = 0
LibraryFlags = 8
LCID = 0x0

from win32com.client import CoClassBaseClass
# This CoClass is known by the name 'AccClientDocMgr.AccClientDocMgr.1'
class AccClientDocMgr(CoClassBaseClass): # A CoClass
	# AccClientDocMgr Class
	CLSID = IID('{FC48CC30-4F3E-4FA1-803B-AD0E196A83B1}')
	coclass_sources = [
	]
	coclass_interfaces = [
	]

# This CoClass is known by the name 'AccDictionary.AccDictionary.1'
class AccDictionary(CoClassBaseClass): # A CoClass
	# AccDictionary Class
	CLSID = IID('{6572EE16-5FE5-4331-BB6D-76A49C56E423}')
	coclass_sources = [
	]
	coclass_interfaces = [
	]

# This CoClass is known by the name 'AccServerDocMgr.AccServerDocMgr.1'
class AccServerDocMgr(CoClassBaseClass): # A CoClass
	# AccServerDocMgr Class
	CLSID = IID('{6089A37E-EB8A-482D-BD6F-F9F46904D16D}')
	coclass_sources = [
	]
	coclass_interfaces = [
	]

class AccStore(CoClassBaseClass): # A CoClass
	# AccStore Class
	CLSID = IID('{5440837F-4BFF-4AE5-A1B1-7722ECC6332A}')
	coclass_sources = [
	]
	coclass_interfaces = [
	]

# This CoClass is known by the name 'DocWrap.DocWrap.1'
class DocWrap(CoClassBaseClass): # A CoClass
	# DocWrap Class
	CLSID = IID('{BF426F7E-7A5E-44D6-830C-A390EA9462A3}')
	coclass_sources = [
	]
	coclass_interfaces = [
	]

class MSAAControl(CoClassBaseClass): # A CoClass
	# MSAAControl Class
	CLSID = IID('{08CD963F-7A3E-4F5C-9BD8-D692BB043C5B}')
	coclass_sources = [
	]
	coclass_interfaces = [
	]

IAccClientDocMgr_vtables_dispatch_ = 0
IAccClientDocMgr_vtables_ = [
	(( 'GetDocuments' , 'enumUnknown' , ), 1610678272, (1610678272, (), [ (16397, 2, None, "IID('{00000100-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 0 , 24 , (3, 0, None, None) , 0 , )),
	(( 'LookupByHWND' , 'hWnd' , 'riid' , 'ppunk' , ), 1610678273, (1610678273, (), [ 
			 (36, 1, None, None) , (36, 1, None, None) , (16397, 2, None, None) , ], 1 , 1 , 4 , 0 , 32 , (3, 0, None, None) , 0 , )),
	(( 'LookupByPoint' , 'pt' , 'riid' , 'ppunk' , ), 1610678274, (1610678274, (), [ 
			 (36, 1, None, None) , (36, 1, None, None) , (16397, 2, None, None) , ], 1 , 1 , 4 , 0 , 40 , (3, 0, None, None) , 0 , )),
	(( 'GetFocused' , 'riid' , 'ppunk' , ), 1610678275, (1610678275, (), [ (36, 1, None, None) , 
			 (16397, 2, None, None) , ], 1 , 1 , 4 , 0 , 48 , (3, 0, None, None) , 0 , )),
]

IAccDictionary_vtables_dispatch_ = 0
IAccDictionary_vtables_ = [
	(( 'GetLocalizedString' , 'Term' , 'lcid' , 'pResult' , 'plcid' , 
			 ), 1610678272, (1610678272, (), [ (36, 1, None, None) , (19, 1, None, None) , (16392, 2, None, None) , (16403, 2, None, None) , ], 1 , 1 , 4 , 0 , 24 , (3, 0, None, None) , 0 , )),
	(( 'GetParentTerm' , 'Term' , 'pParentTerm' , ), 1610678273, (1610678273, (), [ (36, 1, None, None) , 
			 (36, 2, None, None) , ], 1 , 1 , 4 , 0 , 32 , (3, 0, None, None) , 0 , )),
	(( 'GetMnemonicString' , 'Term' , 'pResult' , ), 1610678274, (1610678274, (), [ (36, 1, None, None) , 
			 (16392, 2, None, None) , ], 1 , 1 , 4 , 0 , 40 , (3, 0, None, None) , 0 , )),
	(( 'LookupMnemonicTerm' , 'bstrMnemonic' , 'pTerm' , ), 1610678275, (1610678275, (), [ (8, 1, None, None) , 
			 (36, 2, None, None) , ], 1 , 1 , 4 , 0 , 48 , (3, 0, None, None) , 0 , )),
	(( 'ConvertValueToString' , 'Term' , 'lcid' , 'varValue' , 'pbstrResult' , 
			 'plcid' , ), 1610678276, (1610678276, (), [ (36, 1, None, None) , (19, 1, None, None) , (12, 1, None, None) , 
			 (16392, 2, None, None) , (16403, 2, None, None) , ], 1 , 1 , 4 , 0 , 56 , (3, 0, None, None) , 0 , )),
]

IAccServerDocMgr_vtables_dispatch_ = 0
IAccServerDocMgr_vtables_ = [
	(( 'NewDocument' , 'riid' , 'punk' , ), 1610678272, (1610678272, (), [ (36, 1, None, None) , 
			 (13, 1, None, None) , ], 1 , 1 , 4 , 0 , 24 , (3, 0, None, None) , 0 , )),
	(( 'RevokeDocument' , 'punk' , ), 1610678273, (1610678273, (), [ (13, 1, None, None) , ], 1 , 1 , 4 , 0 , 32 , (3, 0, None, None) , 0 , )),
	(( 'OnDocumentFocus' , 'punk' , ), 1610678274, (1610678274, (), [ (13, 1, None, None) , ], 1 , 1 , 4 , 0 , 40 , (3, 0, None, None) , 0 , )),
]

IAccStore_vtables_dispatch_ = 0
IAccStore_vtables_ = [
	(( 'Register' , 'riid' , 'punk' , ), 1610678272, (1610678272, (), [ (36, 1, None, None) , 
			 (13, 1, None, None) , ], 1 , 1 , 4 , 0 , 24 , (3, 0, None, None) , 0 , )),
	(( 'Unregister' , 'punk' , ), 1610678273, (1610678273, (), [ (13, 1, None, None) , ], 1 , 1 , 4 , 0 , 32 , (3, 0, None, None) , 0 , )),
	(( 'GetDocuments' , 'enumUnknown' , ), 1610678274, (1610678274, (), [ (16397, 2, None, "IID('{00000100-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 0 , 40 , (3, 0, None, None) , 0 , )),
	(( 'LookupByHWND' , 'hWnd' , 'riid' , 'ppunk' , ), 1610678275, (1610678275, (), [ 
			 (36, 1, None, None) , (36, 1, None, None) , (16397, 2, None, None) , ], 1 , 1 , 4 , 0 , 48 , (3, 0, None, None) , 0 , )),
	(( 'LookupByPoint' , 'pt' , 'riid' , 'ppunk' , ), 1610678276, (1610678276, (), [ 
			 (36, 1, None, None) , (36, 1, None, None) , (16397, 2, None, None) , ], 1 , 1 , 4 , 0 , 56 , (3, 0, None, None) , 0 , )),
	(( 'OnDocumentFocus' , 'punk' , ), 1610678277, (1610678277, (), [ (13, 1, None, None) , ], 1 , 1 , 4 , 0 , 64 , (3, 0, None, None) , 0 , )),
	(( 'GetFocused' , 'riid' , 'ppunk' , ), 1610678278, (1610678278, (), [ (36, 1, None, None) , 
			 (16397, 2, None, None) , ], 1 , 1 , 4 , 0 , 72 , (3, 0, None, None) , 0 , )),
]

IDocWrap_vtables_dispatch_ = 0
IDocWrap_vtables_ = [
	(( 'SetDoc' , 'riid' , 'punk' , ), 1610678272, (1610678272, (), [ (36, 1, None, None) , 
			 (13, 1, None, None) , ], 1 , 1 , 4 , 0 , 24 , (3, 0, None, None) , 0 , )),
	(( 'GetWrappedDoc' , 'riid' , 'ppunk' , ), 1610678273, (1610678273, (), [ (36, 1, None, None) , 
			 (16397, 2, None, None) , ], 1 , 1 , 4 , 0 , 32 , (3, 0, None, None) , 0 , )),
]

IEnumUnknown_vtables_dispatch_ = 0
IEnumUnknown_vtables_ = [
	(( 'RemoteNext' , 'celt' , 'rgelt' , 'pceltFetched' , ), 1610678272, (1610678272, (), [ 
			 (19, 1, None, None) , (16397, 2, None, None) , (16403, 2, None, None) , ], 1 , 1 , 4 , 0 , 24 , (3, 0, None, None) , 0 , )),
	(( 'Skip' , 'celt' , ), 1610678273, (1610678273, (), [ (19, 1, None, None) , ], 1 , 1 , 4 , 0 , 32 , (3, 0, None, None) , 0 , )),
	(( 'Reset' , ), 1610678274, (1610678274, (), [ ], 1 , 1 , 4 , 0 , 40 , (3, 0, None, None) , 0 , )),
	(( 'Clone' , 'ppenum' , ), 1610678275, (1610678275, (), [ (16397, 2, None, "IID('{00000100-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 0 , 48 , (3, 0, None, None) , 0 , )),
]

ITfMSAAControl_vtables_dispatch_ = 0
ITfMSAAControl_vtables_ = [
	(( 'SystemEnableMSAA' , ), 1610678272, (1610678272, (), [ ], 1 , 1 , 4 , 0 , 24 , (3, 0, None, None) , 0 , )),
	(( 'SystemDisableMSAA' , ), 1610678273, (1610678273, (), [ ], 1 , 1 , 4 , 0 , 32 , (3, 0, None, None) , 0 , )),
]

RecordMap = {
	###'tagPOINT': '{00000000-0000-0000-0000-000000000000}', # Record disabled because it doesn't have a non-null GUID
}

CLSIDToClassMap = {
	'{08CD963F-7A3E-4F5C-9BD8-D692BB043C5B}' : MSAAControl,
	'{5440837F-4BFF-4AE5-A1B1-7722ECC6332A}' : AccStore,
	'{6572EE16-5FE5-4331-BB6D-76A49C56E423}' : AccDictionary,
	'{6089A37E-EB8A-482D-BD6F-F9F46904D16D}' : AccServerDocMgr,
	'{FC48CC30-4F3E-4FA1-803B-AD0E196A83B1}' : AccClientDocMgr,
	'{BF426F7E-7A5E-44D6-830C-A390EA9462A3}' : DocWrap,
}
CLSIDToPackageMap = {}
win32com.client.CLSIDToClass.RegisterCLSIDsFromDict( CLSIDToClassMap )
VTablesToPackageMap = {}
VTablesToClassMap = {
	'{B5F8FB3B-393F-4F7C-84CB-504924C2705A}' : 'ITfMSAAControl',
	'{E2CD4A63-2B72-4D48-B739-95E4765195BA}' : 'IAccStore',
	'{00000100-0000-0000-C000-000000000046}' : 'IEnumUnknown',
	'{1DC4CB5F-D737-474D-ADE9-5CCFC9BC1CC9}' : 'IAccDictionary',
	'{AD7C73CF-6DD5-4855-ABC2-B04BAD5B9153}' : 'IAccServerDocMgr',
	'{4C896039-7B6D-49E6-A8C1-45116A98292B}' : 'IAccClientDocMgr',
	'{DCD285FE-0BE0-43BD-99C9-AAAEC513C555}' : 'IDocWrap',
}


NamesToIIDMap = {
	'ITfMSAAControl' : '{B5F8FB3B-393F-4F7C-84CB-504924C2705A}',
	'IAccStore' : '{E2CD4A63-2B72-4D48-B739-95E4765195BA}',
	'IEnumUnknown' : '{00000100-0000-0000-C000-000000000046}',
	'IAccDictionary' : '{1DC4CB5F-D737-474D-ADE9-5CCFC9BC1CC9}',
	'IAccServerDocMgr' : '{AD7C73CF-6DD5-4855-ABC2-B04BAD5B9153}',
	'IAccClientDocMgr' : '{4C896039-7B6D-49E6-A8C1-45116A98292B}',
	'IDocWrap' : '{DCD285FE-0BE0-43BD-99C9-AAAEC513C555}',
}


