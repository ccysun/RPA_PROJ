# -*- coding: mbcs -*-
# Created by makepy.py version 0.5.01
# By python version 3.9.1 (tags/v3.9.1:1e5d33e, Dec  7 2020, 17:08:21) [MSC v.1927 64 bit (AMD64)]
# From type library 'HwpCtrl.ocx'
# On Mon Jun  7 15:14:56 2021
'HwpCtrl ActiveX Control module'
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

CLSID = IID('{86E6E883-3879-4514-A25B-81023981DA19}')
MajorVersion = 1
MinorVersion = 2
LibraryFlags = 10
LCID = 0x0

from win32com.client import DispatchBaseClass
class DHwpAction(DispatchBaseClass):
    CLSID = IID('{5227AA3C-8332-4F0A-9B82-D7D89A5CF395}')
    coclass_clsid = IID('{39F4EB62-6155-446D-B210-1F010EED7062}')

    def CreateSet(self):
        ret = self._oleobj_.InvokeTypes(3, LCID, 1, (9, 0), (),)
        if ret is not None:
            ret = Dispatch(ret, 'CreateSet', None)
        return ret

    def Execute(self, param=defaultNamedNotOptArg):
        return self._oleobj_.InvokeTypes(4, LCID, 1, (3, 0), ((9, 0),),param
            )

    def GetDefault(self, param=defaultNamedNotOptArg):
        return self._oleobj_.InvokeTypes(5, LCID, 1, (3, 0), ((9, 0),),param
            )

    def PopupDialog(self, param=defaultNamedNotOptArg):
        return self._oleobj_.InvokeTypes(6, LCID, 1, (3, 0), ((9, 0),),param
            )

    def Run(self):
        return self._oleobj_.InvokeTypes(7, LCID, 1, (24, 0), (),)

    _prop_map_get_ = {
        "actid": (1, 2, (8, 0), (), "actid", None),
        "setid": (2, 2, (8, 0), (), "setid", None),
    }
    _prop_map_put_ = {
        "actid" : ((1, LCID, 4, 0),()),
        "setid" : ((2, LCID, 4, 0),()),
    }
    def __iter__(self):
        "Return a Python iterator for this object"
        try:
            ob = self._oleobj_.InvokeTypes(-4,LCID,3,(13, 10),())
        except pythoncom.error:
            raise TypeError("This object does not support enumeration")
        return win32com.client.util.Iterator(ob, None)

class DHwpCtrlCode(DispatchBaseClass):
    CLSID = IID('{39E84D56-02DA-4AB0-B0C9-4F8AD94E3DF6}')
    coclass_clsid = IID('{C2452DD4-E559-4BF8-B6BF-ACCC3540A02F}')

    def GetAnchorPos(self, type=defaultNamedNotOptArg):
        ret = self._oleobj_.InvokeTypes(15001, LCID, 1, (9, 0), ((3, 0),),type
            )
        if ret is not None:
            ret = Dispatch(ret, 'GetAnchorPos', None)
        return ret

    _prop_map_get_ = {
        "CtrlCh": (1, 2, (2, 0), (), "CtrlCh", None),
        "Next": (3, 2, (9, 0), (), "Next", None),
        "Prev": (4, 2, (9, 0), (), "Prev", None),
        "Properties": (5, 2, (9, 0), (), "Properties", None),
        "UserDesc": (6, 2, (8, 0), (), "UserDesc", None),
        "ctrlid": (2, 2, (8, 0), (), "ctrlid", None),
    }
    _prop_map_put_ = {
        "CtrlCh" : ((1, LCID, 4, 0),()),
        "Next" : ((3, LCID, 4, 0),()),
        "Prev" : ((4, LCID, 4, 0),()),
        "Properties" : ((5, LCID, 4, 0),()),
        "UserDesc" : ((6, LCID, 4, 0),()),
        "ctrlid" : ((2, LCID, 4, 0),()),
    }
    def __iter__(self):
        "Return a Python iterator for this object"
        try:
            ob = self._oleobj_.InvokeTypes(-4,LCID,3,(13, 10),())
        except pythoncom.error:
            raise TypeError("This object does not support enumeration")
        return win32com.client.util.Iterator(ob, None)

class DHwpMenu(DispatchBaseClass):
    CLSID = IID('{982B2F01-9EAB-4783-AE1D-EA7BAFEFCE2F}')
    coclass_clsid = IID('{C9153331-791D-4AF5-AD10-17AC64012A5D}')

    def AppendMenu(self, strActionID=defaultNamedNotOptArg, strMenuName=defaultNamedNotOptArg):
        return self._oleobj_.InvokeTypes(15001, LCID, 1, (3, 0), ((8, 0), (8, 0)),strActionID
            , strMenuName)

    def RemoveMenu(self, strActionID=defaultNamedNotOptArg):
        return self._oleobj_.InvokeTypes(15002, LCID, 1, (3, 0), ((8, 0),),strActionID
            )

    _prop_map_get_ = {
    }
    _prop_map_put_ = {
    }
    def __iter__(self):
        "Return a Python iterator for this object"
        try:
            ob = self._oleobj_.InvokeTypes(-4,LCID,3,(13, 10),())
        except pythoncom.error:
            raise TypeError("This object does not support enumeration")
        return win32com.client.util.Iterator(ob, None)

class DHwpParameterArray(DispatchBaseClass):
    CLSID = IID('{E35C1FFB-BAD5-44B7-A8F2-7B598123C2FC}')
    coclass_clsid = IID('{30517C18-C328-4C1E-AE70-4BE894335FB3}')

    def Clone(self):
        ret = self._oleobj_.InvokeTypes(3, LCID, 1, (9, 0), (),)
        if ret is not None:
            ret = Dispatch(ret, 'Clone', None)
        return ret

    def Copy(self, srcarray=defaultNamedNotOptArg):
        return self._oleobj_.InvokeTypes(4, LCID, 1, (24, 0), ((9, 0),),srcarray
            )

    def Item(self, index=defaultNamedNotOptArg):
        return self._ApplyTypes_(5, 1, (12, 0), ((3, 0),), 'Item', None,index
            )

    def SetItem(self, index=defaultNamedNotOptArg, value=defaultNamedNotOptArg):
        return self._oleobj_.InvokeTypes(6, LCID, 1, (24, 0), ((3, 0), (12, 0)),index
            , value)

    _prop_map_get_ = {
        "Count": (1, 2, (3, 0), (), "Count", None),
        "IsSet": (2, 2, (11, 0), (), "IsSet", None),
    }
    _prop_map_put_ = {
        "Count" : ((1, LCID, 4, 0),()),
        "IsSet" : ((2, LCID, 4, 0),()),
    }
    def __iter__(self):
        "Return a Python iterator for this object"
        try:
            ob = self._oleobj_.InvokeTypes(-4,LCID,3,(13, 10),())
        except pythoncom.error:
            raise TypeError("This object does not support enumeration")
        return win32com.client.util.Iterator(ob, None)
    #This class has Item property/method which allows indexed access with the object[key] syntax.
    #Some objects will accept a string or other type of key in addition to integers.
    #Note that many Office objects do not use zero-based indexing.
    def __getitem__(self, key):
        return self._get_good_object_(self._oleobj_.Invoke(*(5, LCID, 1, 1, key)), "Item", None)
    #This class has Count() property - allow len(ob) to provide this
    def __len__(self):
        return self._ApplyTypes_(*(1, 2, (3, 0), (), "Count", None))
    #This class has a __len__ - this is needed so 'if object:' always returns TRUE.
    def __nonzero__(self):
        return True

class DHwpParameterSet(DispatchBaseClass):
    CLSID = IID('{E5BBBA9F-732F-450D-8433-1B7AA4F60503}')
    coclass_clsid = IID('{FA18D456-0798-4AD0-AC3A-82A3A7DB1519}')

    def Clone(self):
        ret = self._oleobj_.InvokeTypes(4, LCID, 1, (9, 0), (),)
        if ret is not None:
            ret = Dispatch(ret, 'Clone', None)
        return ret

    def CreateItemArray(self, itemid=defaultNamedNotOptArg, Count=defaultNamedNotOptArg):
        ret = self._oleobj_.InvokeTypes(5, LCID, 1, (9, 0), ((8, 0), (3, 0)),itemid
            , Count)
        if ret is not None:
            ret = Dispatch(ret, 'CreateItemArray', None)
        return ret

    def CreateItemSet(self, itemid=defaultNamedNotOptArg, setid=defaultNamedNotOptArg):
        ret = self._oleobj_.InvokeTypes(6, LCID, 1, (9, 0), ((8, 0), (8, 0)),itemid
            , setid)
        if ret is not None:
            ret = Dispatch(ret, 'CreateItemSet', None)
        return ret

    def GetIntersection(self, srcset=defaultNamedNotOptArg):
        return self._oleobj_.InvokeTypes(7, LCID, 1, (24, 0), ((9, 0),),srcset
            )

    def IsEquivalent(self, srcset=defaultNamedNotOptArg):
        return self._oleobj_.InvokeTypes(8, LCID, 1, (11, 0), ((9, 0),),srcset
            )

    def Item(self, itemid=defaultNamedNotOptArg):
        return self._ApplyTypes_(9, 1, (12, 0), ((8, 0),), 'Item', None,itemid
            )

    def ItemExist(self, itemid=defaultNamedNotOptArg):
        return self._oleobj_.InvokeTypes(10, LCID, 1, (11, 0), ((8, 0),),itemid
            )

    def Merge(self, srcset=defaultNamedNotOptArg):
        return self._oleobj_.InvokeTypes(11, LCID, 1, (24, 0), ((9, 0),),srcset
            )

    def RemoveAll(self, setid=defaultNamedNotOptArg):
        return self._oleobj_.InvokeTypes(12, LCID, 1, (24, 0), ((8, 0),),setid
            )

    def RemoveItem(self, itemid=defaultNamedNotOptArg):
        return self._oleobj_.InvokeTypes(13, LCID, 1, (24, 0), ((8, 0),),itemid
            )

    def SetItem(self, itemid=defaultNamedNotOptArg, value=defaultNamedNotOptArg):
        return self._oleobj_.InvokeTypes(14, LCID, 1, (24, 0), ((8, 0), (12, 0)),itemid
            , value)

    _prop_map_get_ = {
        "Count": (1, 2, (3, 0), (), "Count", None),
        "IsSet": (2, 2, (11, 0), (), "IsSet", None),
        "setid": (3, 2, (8, 0), (), "setid", None),
    }
    _prop_map_put_ = {
        "Count" : ((1, LCID, 4, 0),()),
        "IsSet" : ((2, LCID, 4, 0),()),
        "setid" : ((3, LCID, 4, 0),()),
    }
    def __iter__(self):
        "Return a Python iterator for this object"
        try:
            ob = self._oleobj_.InvokeTypes(-4,LCID,3,(13, 10),())
        except pythoncom.error:
            raise TypeError("This object does not support enumeration")
        return win32com.client.util.Iterator(ob, None)
    #This class has Item property/method which allows indexed access with the object[key] syntax.
    #Some objects will accept a string or other type of key in addition to integers.
    #Note that many Office objects do not use zero-based indexing.
    def __getitem__(self, key):
        return self._get_good_object_(self._oleobj_.Invoke(*(9, LCID, 1, 1, key)), "Item", None)
    #This class has Count() property - allow len(ob) to provide this
    def __len__(self):
        return self._ApplyTypes_(*(1, 2, (3, 0), (), "Count", None))
    #This class has a __len__ - this is needed so 'if object:' always returns TRUE.
    def __nonzero__(self):
        return True

class _DHwpCtrl(DispatchBaseClass):
    'Dispatch interface for HwpCtrl Control'
    CLSID = IID('{377C0BC8-E22C-45ED-851A-FBA0208DDC23}')
    coclass_clsid = IID('{BD9C32DE-3155-4691-8972-097D53B10052}')

    def AboutBox(self):
        return self._oleobj_.InvokeTypes(-552, LCID, 1, (24, 0), (),)

    def Clear(self, option=defaultNamedOptArg):
        return self._oleobj_.InvokeTypes(16, LCID, 1, (24, 0), ((12, 16),),option
            )

    def ConvertPUAHangulToUnicode(self, reverse=defaultNamedOptArg):
        return self._oleobj_.InvokeTypes(16000, LCID, 1, (3, 0), ((12, 16),),reverse
            )

    def CreateAction(self, actid=defaultNamedNotOptArg):
        ret = self._oleobj_.InvokeTypes(17, LCID, 1, (9, 0), ((8, 0),),actid
            )
        if ret is not None:
            ret = Dispatch(ret, 'CreateAction', None)
        return ret

    def CreateField(self, direction=defaultNamedNotOptArg, memo=defaultNamedOptArg, name=defaultNamedOptArg):
        return self._oleobj_.InvokeTypes(39, LCID, 1, (11, 0), ((8, 0), (12, 16), (12, 16)),direction
            , memo, name)

    def CreatePageImage(self, Path=defaultNamedNotOptArg, pgno=defaultNamedOptArg, resolution=defaultNamedOptArg, depth=defaultNamedOptArg
            , format=defaultNamedOptArg):
        return self._oleobj_.InvokeTypes(18, LCID, 1, (11, 0), ((8, 0), (12, 16), (12, 16), (12, 16), (12, 16)),Path
            , pgno, resolution, depth, format)

    def CreatePageImageEx(self, Path=defaultNamedNotOptArg, pgno=defaultNamedOptArg, resolution=defaultNamedOptArg, depth=defaultNamedOptArg
            , format=defaultNamedOptArg, options=defaultNamedOptArg):
        return self._oleobj_.InvokeTypes(15101, LCID, 1, (11, 0), ((8, 0), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),Path
            , pgno, resolution, depth, format, options
            )

    def CreateSet(self, setid=defaultNamedNotOptArg):
        ret = self._oleobj_.InvokeTypes(19, LCID, 1, (9, 0), ((8, 0),),setid
            )
        if ret is not None:
            ret = Dispatch(ret, 'CreateSet', None)
        return ret

    def DeleteCtrl(self, ctrl=defaultNamedNotOptArg):
        return self._oleobj_.InvokeTypes(45, LCID, 1, (11, 0), ((9, 0),),ctrl
            )

    def ExportStyle(self, styleset=defaultNamedNotOptArg):
        return self._oleobj_.InvokeTypes(15079, LCID, 1, (11, 0), ((9, 0),),styleset
            )

    def FieldExist(self, field=defaultNamedNotOptArg):
        return self._oleobj_.InvokeTypes(20, LCID, 1, (11, 0), ((8, 0),),field
            )

    def FindPrivateInfo(self, PrivateType=defaultNamedNotOptArg, PrivateString=defaultNamedOptArg):
        return self._oleobj_.InvokeTypes(15089, LCID, 1, (3, 0), ((3, 0), (12, 16)),PrivateType
            , PrivateString)

    def GetActionCmdUIStatus(self, actid=defaultNamedNotOptArg, bWithKey=defaultNamedNotOptArg, bEnabled=defaultNamedNotOptArg, bChecked=defaultNamedNotOptArg
            , bRadio=defaultNamedNotOptArg, szText=defaultNamedNotOptArg):
        return self._oleobj_.InvokeTypes(55, LCID, 1, (11, 0), ((8, 0), (3, 0), (16387, 0), (16387, 0), (16387, 0), (16392, 0)),actid
            , bWithKey, bEnabled, bChecked, bRadio, szText
            )

    def GetBinDataPath(self, binid=defaultNamedNotOptArg):
        # Result is a Unicode object
        return self._oleobj_.InvokeTypes(15081, LCID, 1, (8, 0), ((3, 0),),binid
            )

    def GetContextMenu(self, contextID=defaultNamedNotOptArg):
        ret = self._oleobj_.InvokeTypes(15086, LCID, 1, (9, 0), ((8, 0),),contextID
            )
        if ret is not None:
            ret = Dispatch(ret, 'GetContextMenu', None)
        return ret

    def GetCtrlHorizontalOffset(self, ctrl=defaultNamedNotOptArg, relTo=defaultNamedNotOptArg):
        return self._oleobj_.InvokeTypes(15091, LCID, 1, (3, 0), ((9, 0), (3, 0)),ctrl
            , relTo)

    def GetCtrlVerticalOffset(self, ctrl=defaultNamedNotOptArg, relTo=defaultNamedNotOptArg):
        return self._oleobj_.InvokeTypes(15092, LCID, 1, (3, 0), ((9, 0), (3, 0)),ctrl
            , relTo)

    def GetCurFieldName(self, option=defaultNamedOptArg):
        # Result is a Unicode object
        return self._oleobj_.InvokeTypes(43, LCID, 1, (8, 0), ((12, 16),),option
            )

    def GetFieldList(self, number=defaultNamedOptArg, option=defaultNamedOptArg):
        # Result is a Unicode object
        return self._oleobj_.InvokeTypes(21, LCID, 1, (8, 0), ((12, 16), (12, 16)),number
            , option)

    def GetFieldText(self, fieldlist=defaultNamedNotOptArg):
        # Result is a Unicode object
        return self._oleobj_.InvokeTypes(22, LCID, 1, (8, 0), ((8, 0),),fieldlist
            )

    def GetFileInfo(self, FileName=defaultNamedNotOptArg):
        ret = self._oleobj_.InvokeTypes(67, LCID, 1, (9, 0), ((8, 0),),FileName
            )
        if ret is not None:
            ret = Dispatch(ret, 'GetFileInfo', None)
        return ret

    def GetFilterList(self, szfilterlist=defaultNamedNotOptArg, flags=defaultNamedNotOptArg):
        return self._oleobj_.InvokeTypes(51, LCID, 1, (11, 0), ((16392, 0), (12, 0)),szfilterlist
            , flags)

    def GetFormObjectAttr(self, formname=defaultNamedNotOptArg, attrname=defaultNamedNotOptArg):
        return self._ApplyTypes_(15060, 1, (12, 0), ((8, 0), (8, 0)), 'GetFormObjectAttr', None,formname
            , attrname)

    def GetHeadingString(self):
        # Result is a Unicode object
        return self._oleobj_.InvokeTypes(15085, LCID, 1, (8, 0), (),)

    def GetMessageBoxMode(self):
        return self._oleobj_.InvokeTypes(15077, LCID, 1, (3, 0), (),)

    def GetMessageSet(self):
        ret = self._oleobj_.InvokeTypes(74, LCID, 1, (9, 0), (),)
        if ret is not None:
            ret = Dispatch(ret, 'GetMessageSet', None)
        return ret

    def GetMousePos(self, Xrelto=defaultNamedNotOptArg, Yrelto=defaultNamedNotOptArg):
        ret = self._oleobj_.InvokeTypes(65, LCID, 1, (9, 0), ((3, 0), (3, 0)),Xrelto
            , Yrelto)
        if ret is not None:
            ret = Dispatch(ret, 'GetMousePos', None)
        return ret

    def GetPageText(self, pgno=defaultNamedNotOptArg, option=defaultNamedOptArg):
        # Result is a Unicode object
        return self._oleobj_.InvokeTypes(15075, LCID, 1, (8, 0), ((3, 0), (12, 16)),pgno
            , option)

    def GetPos(self, list=defaultNamedNotOptArg, para=defaultNamedNotOptArg, pos=defaultNamedNotOptArg):
        return self._oleobj_.InvokeTypes(52, LCID, 1, (24, 0), ((16387, 0), (16387, 0), (16387, 0)),list
            , para, pos)

    def GetPosBySet(self):
        ret = self._oleobj_.InvokeTypes(58, LCID, 1, (9, 0), (),)
        if ret is not None:
            ret = Dispatch(ret, 'GetPosBySet', None)
        return ret

    def GetScriptSource(self, filepath=defaultNamedNotOptArg):
        # Result is a Unicode object
        return self._oleobj_.InvokeTypes(15071, LCID, 1, (8, 0), ((8, 0),),filepath
            )

    def GetSelectedPos(self, slist=defaultNamedNotOptArg, spara=defaultNamedNotOptArg, spos=defaultNamedNotOptArg, elist=defaultNamedNotOptArg
            , epara=defaultNamedNotOptArg, epos=defaultNamedNotOptArg):
        return self._oleobj_.InvokeTypes(15063, LCID, 1, (11, 0), ((16387, 0), (16387, 0), (16387, 0), (16387, 0), (16387, 0), (16387, 0)),slist
            , spara, spos, elist, epara, epos
            )

    def GetSelectedPosBySet(self, sset=defaultNamedNotOptArg, eset=defaultNamedNotOptArg):
        return self._oleobj_.InvokeTypes(15064, LCID, 1, (11, 0), ((9, 0), (9, 0)),sset
            , eset)

    def GetTableCellAddr(self, type=defaultNamedNotOptArg):
        return self._oleobj_.InvokeTypes(15065, LCID, 1, (3, 0), ((3, 0),),type
            )

    def GetText(self, text=defaultNamedNotOptArg):
        return self._oleobj_.InvokeTypes(47, LCID, 1, (3, 0), ((16392, 0),),text
            )

    def GetTextBySet(self, text=defaultNamedNotOptArg):
        return self._oleobj_.InvokeTypes(60, LCID, 1, (3, 0), ((9, 0),),text
            )

    def GetTextFile(self, format=defaultNamedNotOptArg, option=defaultNamedNotOptArg):
        return self._ApplyTypes_(63, 1, (12, 0), ((8, 0), (8, 0)), 'GetTextFile', None,format
            , option)

    def GetVersionHistoryCount(self):
        return self._oleobj_.InvokeTypes(15070, LCID, 1, (3, 0), (),)

    def GetVersionInfo(self, index=defaultNamedNotOptArg):
        ret = self._oleobj_.InvokeTypes(15069, LCID, 1, (9, 0), ((3, 0),),index
            )
        if ret is not None:
            ret = Dispatch(ret, 'GetVersionInfo', None)
        return ret

    def GetViewStatus(self, nType=defaultNamedNotOptArg):
        ret = self._oleobj_.InvokeTypes(75, LCID, 1, (9, 0), ((3, 0),),nType
            )
        if ret is not None:
            ret = Dispatch(ret, 'GetViewStatus', None)
        return ret

    def ImportStyle(self, styleset=defaultNamedNotOptArg):
        return self._oleobj_.InvokeTypes(15080, LCID, 1, (11, 0), ((9, 0),),styleset
            )

    def InitScan(self, option=defaultNamedOptArg, range=defaultNamedOptArg, spara=defaultNamedOptArg, spos=defaultNamedOptArg
            , epara=defaultNamedOptArg, epos=defaultNamedOptArg):
        return self._oleobj_.InvokeTypes(46, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),option
            , range, spara, spos, epara, epos
            )

    def Insert(self, Path=defaultNamedNotOptArg, format=defaultNamedOptArg, arg=defaultNamedOptArg):
        return self._oleobj_.InvokeTypes(23, LCID, 1, (11, 0), ((8, 0), (12, 16), (12, 16)),Path
            , format, arg)

    def InsertBackgroundPicture(self, bordertype=defaultNamedNotOptArg, Path=defaultNamedNotOptArg, embedded=defaultNamedOptArg, filloption=defaultNamedOptArg
            , watermark=defaultNamedOptArg, effect=defaultNamedOptArg, brightness=defaultNamedOptArg, contrast=defaultNamedOptArg):
        return self._oleobj_.InvokeTypes(61, LCID, 1, (11, 0), ((8, 0), (8, 0), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),bordertype
            , Path, embedded, filloption, watermark, effect
            , brightness, contrast)

    def InsertCtrl(self, ctrlid=defaultNamedNotOptArg, initparam=defaultNamedOptArg):
        ret = self._oleobj_.InvokeTypes(24, LCID, 1, (9, 0), ((8, 0), (12, 16)),ctrlid
            , initparam)
        if ret is not None:
            ret = Dispatch(ret, 'InsertCtrl', None)
        return ret

    def InsertDocument(self, szFileName=defaultNamedNotOptArg):
        return self._oleobj_.InvokeTypes(34, LCID, 1, (11, 0), ((8, 0),),szFileName
            )

    def InsertMenu(self, menuidx=defaultNamedNotOptArg, menustr=defaultNamedNotOptArg, actionstr=defaultNamedNotOptArg, menutype=defaultNamedNotOptArg):
        return self._oleobj_.InvokeTypes(15073, LCID, 1, (11, 0), ((8, 0), (8, 0), (8, 0), (3, 0)),menuidx
            , menustr, actionstr, menutype)

    def InsertPicture(self, Path=defaultNamedNotOptArg, embedded=defaultNamedOptArg, sizeOption=defaultNamedOptArg, reverse=defaultNamedOptArg
            , watermark=defaultNamedOptArg, effect=defaultNamedOptArg, width=defaultNamedOptArg, height=defaultNamedOptArg):
        ret = self._oleobj_.InvokeTypes(38, LCID, 1, (9, 0), ((8, 0), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16), (12, 16)),Path
            , embedded, sizeOption, reverse, watermark, effect
            , width, height)
        if ret is not None:
            ret = Dispatch(ret, 'InsertPicture', None)
        return ret

    def IsCommandLock(self, actionID=defaultNamedNotOptArg):
        return self._oleobj_.InvokeTypes(50, LCID, 1, (11, 0), ((8, 0),),actionID
            )

    def IsSpellCheckCompleted(self):
        return self._oleobj_.InvokeTypes(15084, LCID, 1, (11, 0), (),)

    def KeyIndicator(self, seccnt=defaultNamedNotOptArg, secno=defaultNamedNotOptArg, prnpageno=defaultNamedNotOptArg, colno=defaultNamedNotOptArg
            , line=defaultNamedNotOptArg, pos=defaultNamedNotOptArg, over=defaultNamedNotOptArg, ctrlname=defaultNamedNotOptArg):
        return self._oleobj_.InvokeTypes(54, LCID, 1, (11, 0), ((16387, 0), (16387, 0), (16387, 0), (16387, 0), (16387, 0), (16387, 0), (16386, 0), (16392, 0)),seccnt
            , secno, prnpageno, colno, line, pos
            , over, ctrlname)

    def LoadState(self, FileName=defaultNamedNotOptArg):
        return self._oleobj_.InvokeTypes(69, LCID, 1, (11, 0), ((8, 0),),FileName
            )

    def LockCommand(self, actionID=defaultNamedNotOptArg, lock=defaultNamedNotOptArg):
        return self._oleobj_.InvokeTypes(49, LCID, 1, (24, 0), ((8, 0), (3, 0)),actionID
            , lock)

    def LunarToSolar(self, lYear=defaultNamedNotOptArg, lMonth=defaultNamedNotOptArg, lDay=defaultNamedNotOptArg, lLeap=defaultNamedNotOptArg
            , sYear=defaultNamedNotOptArg, sMonth=defaultNamedNotOptArg, sDay=defaultNamedNotOptArg):
        return self._oleobj_.InvokeTypes(15095, LCID, 1, (11, 0), ((3, 0), (3, 0), (3, 0), (3, 0), (16387, 0), (16387, 0), (16387, 0)),lYear
            , lMonth, lDay, lLeap, sYear, sMonth
            , sDay)

    def LunarToSolarBySet(self, lYear=defaultNamedNotOptArg, lMonth=defaultNamedNotOptArg, lDay=defaultNamedNotOptArg, lLeap=defaultNamedNotOptArg):
        ret = self._oleobj_.InvokeTypes(15096, LCID, 1, (9, 0), ((3, 0), (3, 0), (3, 0), (3, 0)),lYear
            , lMonth, lDay, lLeap)
        if ret is not None:
            ret = Dispatch(ret, 'LunarToSolarBySet', None)
        return ret

    def MakeDocumentDiff(self, srcFilePath=defaultNamedNotOptArg, tgtFilePath=defaultNamedNotOptArg, srcResultFilePath=defaultNamedNotOptArg, tgtResultFilePath=defaultNamedNotOptArg
            , viewWithMemo=defaultNamedNotOptArg):
        return self._oleobj_.InvokeTypes(15082, LCID, 1, (3, 0), ((8, 0), (8, 0), (8, 0), (8, 0), (3, 0)),srcFilePath
            , tgtFilePath, srcResultFilePath, tgtResultFilePath, viewWithMemo)

    def MakeDocumentMergeDiff(self, srcFilePath=defaultNamedNotOptArg, tgtFilePath=defaultNamedNotOptArg, mergeResultFilePath=defaultNamedNotOptArg, viewWithMemo=defaultNamedNotOptArg):
        return self._oleobj_.InvokeTypes(15083, LCID, 1, (3, 0), ((8, 0), (8, 0), (8, 0), (3, 0)),srcFilePath
            , tgtFilePath, mergeResultFilePath, viewWithMemo)

    def MakeVersionDiffAll(self, filepath=defaultNamedOptArg):
        ret = self._oleobj_.InvokeTypes(15068, LCID, 1, (9, 0), ((12, 16),),filepath
            )
        if ret is not None:
            ret = Dispatch(ret, 'MakeVersionDiffAll', None)
        return ret

    def ModifyFieldProperties(self, field=defaultNamedNotOptArg, remove=defaultNamedNotOptArg, add=defaultNamedNotOptArg):
        return self._oleobj_.InvokeTypes(57, LCID, 1, (3, 0), ((8, 0), (3, 0), (3, 0)),field
            , remove, add)

    def MovePos(self, moveID=defaultNamedOptArg, para=defaultNamedOptArg, pos=defaultNamedOptArg):
        return self._oleobj_.InvokeTypes(41, LCID, 1, (11, 0), ((12, 16), (12, 16), (12, 16)),moveID
            , para, pos)

    def MoveToField(self, field=defaultNamedNotOptArg, text=defaultNamedOptArg, start=defaultNamedOptArg, select=defaultNamedOptArg):
        return self._oleobj_.InvokeTypes(25, LCID, 1, (11, 0), ((8, 0), (12, 16), (12, 16), (12, 16)),field
            , text, start, select)

    def MoveToFieldEx(self, field=defaultNamedNotOptArg, text=defaultNamedOptArg, start=defaultNamedOptArg, select=defaultNamedOptArg):
        return self._oleobj_.InvokeTypes(77, LCID, 1, (11, 0), ((8, 0), (12, 16), (12, 16), (12, 16)),field
            , text, start, select)

    def Open(self, Path=defaultNamedNotOptArg, format=defaultNamedOptArg, arg=defaultNamedOptArg):
        return self._oleobj_.InvokeTypes(26, LCID, 1, (11, 0), ((8, 0), (12, 16), (12, 16)),Path
            , format, arg)

    def OpenDocument(self, szFileName=defaultNamedNotOptArg, szFileType=defaultNamedNotOptArg):
        return self._oleobj_.InvokeTypes(32, LCID, 1, (11, 0), ((8, 0), (8, 0)),szFileName
            , szFileType)

    def PreviewCommand(self, previewmode=defaultNamedNotOptArg):
        return self._oleobj_.InvokeTypes(15062, LCID, 1, (11, 0), ((3, 0),),previewmode
            )

    def PrintDocument(self):
        return self._oleobj_.InvokeTypes(35, LCID, 1, (11, 0), (),)

    def ProtectPrivateInfo(self, PotectingChar=defaultNamedNotOptArg, PrivatePatternType=defaultNamedOptArg):
        return self._oleobj_.InvokeTypes(15090, LCID, 1, (3, 0), ((8, 0), (12, 16)),PotectingChar
            , PrivatePatternType)

    def PutFieldText(self, fieldlist=defaultNamedNotOptArg, textlist=defaultNamedNotOptArg):
        return self._oleobj_.InvokeTypes(27, LCID, 1, (24, 0), ((8, 0), (8, 0)),fieldlist
            , textlist)

    def RegisterModule(self, ModuleType=defaultNamedNotOptArg, ModuleData=defaultNamedNotOptArg):
        return self._oleobj_.InvokeTypes(76, LCID, 1, (11, 0), ((8, 0), (12, 0)),ModuleType
            , ModuleData)

    def RegisterPrivateInfoPattern(self, PrivateType=defaultNamedNotOptArg, PrivatePattern=defaultNamedNotOptArg):
        return self._oleobj_.InvokeTypes(15088, LCID, 1, (3, 0), ((3, 0), (8, 0)),PrivateType
            , PrivatePattern)

    def ReleaseScan(self):
        return self._oleobj_.InvokeTypes(48, LCID, 1, (24, 0), (),)

    def RemoveMenu(self, menuidx=defaultNamedNotOptArg, menutype=defaultNamedNotOptArg):
        return self._oleobj_.InvokeTypes(15074, LCID, 1, (11, 0), ((8, 0), (3, 0)),menuidx
            , menutype)

    def RenameField(self, oldname=defaultNamedNotOptArg, newname=defaultNamedNotOptArg):
        return self._oleobj_.InvokeTypes(28, LCID, 1, (24, 0), ((8, 0), (8, 0)),oldname
            , newname)

    def ReplaceAction(self, OldActionID=defaultNamedNotOptArg, NewActionID=defaultNamedNotOptArg):
        return self._oleobj_.InvokeTypes(70, LCID, 1, (11, 0), ((8, 0), (8, 0)),OldActionID
            , NewActionID)

    def Run(self, actid=defaultNamedNotOptArg):
        return self._oleobj_.InvokeTypes(29, LCID, 1, (24, 0), ((8, 0),),actid
            )

    def RunScriptMacro(self, FunctionName=defaultNamedNotOptArg, uMacroType=defaultNamedNotOptArg, uScriptType=defaultNamedNotOptArg):
        return self._oleobj_.InvokeTypes(15072, LCID, 1, (11, 0), ((8, 0), (3, 0), (3, 0)),FunctionName
            , uMacroType, uScriptType)

    def Save(self, save_if_dirty=defaultNamedOptArg):
        return self._oleobj_.InvokeTypes(30, LCID, 1, (11, 0), ((12, 16),),save_if_dirty
            )

    def SaveAs(self, Path=defaultNamedNotOptArg, format=defaultNamedOptArg, arg=defaultNamedOptArg):
        return self._oleobj_.InvokeTypes(31, LCID, 1, (11, 0), ((8, 0), (12, 16), (12, 16)),Path
            , format, arg)

    def SaveDocument(self, szFileName=defaultNamedNotOptArg, szFileType=defaultNamedNotOptArg):
        return self._oleobj_.InvokeTypes(33, LCID, 1, (11, 0), ((8, 0), (8, 0)),szFileName
            , szFileType)

    def SaveState(self, FileName=defaultNamedNotOptArg):
        # Result is a Unicode object
        return self._oleobj_.InvokeTypes(68, LCID, 1, (8, 0), ((8, 0),),FileName
            )

    def SelectText(self, spara=defaultNamedNotOptArg, spos=defaultNamedNotOptArg, epara=defaultNamedNotOptArg, epos=defaultNamedNotOptArg):
        return self._oleobj_.InvokeTypes(42, LCID, 1, (11, 0), ((3, 0), (3, 0), (3, 0), (3, 0)),spara
            , spos, epara, epos)

    def SetAutoSave(self, FileName=defaultNamedNotOptArg, saveinterval=defaultNamedNotOptArg):
        return self._oleobj_.InvokeTypes(15059, LCID, 1, (11, 0), ((8, 0), (3, 0)),FileName
            , saveinterval)

    def SetBarCodeImage(self, lpImagePath=defaultNamedNotOptArg, pgno=defaultNamedNotOptArg, index=defaultNamedNotOptArg, x=defaultNamedNotOptArg
            , y=defaultNamedNotOptArg, width=defaultNamedOptArg, height=defaultNamedOptArg):
        return self._oleobj_.InvokeTypes(15076, LCID, 1, (11, 0), ((8, 0), (3, 0), (3, 0), (3, 0), (3, 0), (12, 16), (12, 16)),lpImagePath
            , pgno, index, x, y, width
            , height)

    def SetClientName(self, szClient=defaultNamedNotOptArg):
        return self._oleobj_.InvokeTypes(40, LCID, 1, (24, 0), ((8, 0),),szClient
            )

    def SetCurFieldName(self, fieldname=defaultNamedNotOptArg, option=defaultNamedOptArg, direction=defaultNamedOptArg, memo=defaultNamedOptArg):
        return self._oleobj_.InvokeTypes(44, LCID, 1, (11, 0), ((8, 0), (12, 16), (12, 16), (12, 16)),fieldname
            , option, direction, memo)

    def SetFieldViewOption(self, option=defaultNamedNotOptArg):
        return self._oleobj_.InvokeTypes(62, LCID, 1, (3, 0), ((3, 0),),option
            )

    def SetFormObjectAttr(self, formname=defaultNamedNotOptArg, attrname=defaultNamedNotOptArg, value=defaultNamedNotOptArg):
        return self._oleobj_.InvokeTypes(15061, LCID, 1, (11, 0), ((8, 0), (8, 0), (12, 0)),formname
            , attrname, value)

    def SetMessageBoxMode(self, mode=defaultNamedNotOptArg):
        return self._oleobj_.InvokeTypes(15078, LCID, 1, (3, 0), ((3, 0),),mode
            )

    def SetPos(self, list=defaultNamedNotOptArg, para=defaultNamedNotOptArg, pos=defaultNamedNotOptArg):
        return self._oleobj_.InvokeTypes(53, LCID, 1, (11, 0), ((3, 0), (3, 0), (3, 0)),list
            , para, pos)

    def SetPosBySet(self, pos=defaultNamedNotOptArg):
        return self._oleobj_.InvokeTypes(59, LCID, 1, (11, 0), ((9, 0),),pos
            )

    def SetPrivateInfoPassword(self, password=defaultNamedNotOptArg):
        return self._oleobj_.InvokeTypes(15087, LCID, 1, (3, 0), ((8, 0),),password
            )

    def SetTextFile(self, data=defaultNamedNotOptArg, format=defaultNamedNotOptArg, option=defaultNamedNotOptArg):
        return self._oleobj_.InvokeTypes(64, LCID, 1, (3, 0), ((12, 0), (8, 0), (8, 0)),data
            , format, option)

    def SetToolBar(self, lToolBarID=defaultNamedNotOptArg, varID=defaultNamedNotOptArg):
        return self._oleobj_.InvokeTypes(36, LCID, 1, (11, 0), ((3, 0), (12, 0)),lToolBarID
            , varID)

    def ShowCaret(self, bShow=defaultNamedNotOptArg):
        return self._oleobj_.InvokeTypes(15099, LCID, 1, (24, 0), ((3, 0),),bShow
            )

    def ShowHorizontalScroll(self, bShow=defaultNamedNotOptArg):
        return self._oleobj_.InvokeTypes(15098, LCID, 1, (24, 0), ((3, 0),),bShow
            )

    def ShowPageTooltip(self, bShow=defaultNamedNotOptArg):
        return self._oleobj_.InvokeTypes(15100, LCID, 1, (24, 0), ((3, 0),),bShow
            )

    def ShowStatusBar(self, Show=defaultNamedNotOptArg):
        return self._oleobj_.InvokeTypes(66, LCID, 1, (11, 0), ((3, 0),),Show
            )

    def ShowToolBar(self, bShow=defaultNamedNotOptArg):
        return self._oleobj_.InvokeTypes(37, LCID, 1, (11, 0), ((3, 0),),bShow
            )

    def ShowVerticalScroll(self, bShow=defaultNamedNotOptArg):
        return self._oleobj_.InvokeTypes(15097, LCID, 1, (24, 0), ((3, 0),),bShow
            )

    def SolarToLunar(self, sYear=defaultNamedNotOptArg, sMonth=defaultNamedNotOptArg, sDay=defaultNamedNotOptArg, lYear=defaultNamedNotOptArg
            , lMonth=defaultNamedNotOptArg, lDay=defaultNamedNotOptArg, lLeap=defaultNamedNotOptArg):
        return self._oleobj_.InvokeTypes(15093, LCID, 1, (11, 0), ((3, 0), (3, 0), (3, 0), (16387, 0), (16387, 0), (16387, 0), (16395, 0)),sYear
            , sMonth, sDay, lYear, lMonth, lDay
            , lLeap)

    def SolarToLunarBySet(self, sYear=defaultNamedNotOptArg, sMonth=defaultNamedNotOptArg, sDay=defaultNamedNotOptArg):
        ret = self._oleobj_.InvokeTypes(15094, LCID, 1, (9, 0), ((3, 0), (3, 0), (3, 0)),sYear
            , sMonth, sDay)
        if ret is not None:
            ret = Dispatch(ret, 'SolarToLunarBySet', None)
        return ret

    def VersionDelete(self, index=defaultNamedNotOptArg):
        return self._oleobj_.InvokeTypes(15067, LCID, 1, (11, 0), ((3, 0),),index
            )

    def VersionSave(self, filepath=defaultNamedNotOptArg, overwirte=defaultNamedNotOptArg, infolock=defaultNamedNotOptArg, writer=defaultNamedOptArg
            , description=defaultNamedOptArg):
        return self._oleobj_.InvokeTypes(15066, LCID, 1, (11, 0), ((8, 0), (3, 0), (3, 0), (12, 16), (12, 16)),filepath
            , overwirte, infolock, writer, description)

    _prop_map_get_ = {
        "AutoShowHideToolBar": (71, 2, (11, 0), (), "AutoShowHideToolBar", None),
        "CellShape": (1, 2, (9, 0), (), "CellShape", None),
        "CharShape": (2, 2, (9, 0), (), "CharShape", None),
        "CurFieldState": (13, 2, (3, 0), (), "CurFieldState", None),
        "CurSelectedCtrl": (123, 2, (9, 0), (), "CurSelectedCtrl", None),
        "EditMode": (56, 2, (3, 0), (), "EditMode", None),
        "EngineProperties": (72, 2, (9, 0), (), "EngineProperties", None),
        "HeadCtrl": (3, 2, (9, 0), (), "HeadCtrl", None),
        "HyperlinkMode": (78, 2, (3, 0), (), "HyperlinkMode", None),
        "IsEmpty": (4, 2, (11, 0), (), "IsEmpty", None),
        "IsModified": (5, 2, (2, 0), (), "IsModified", None),
        "IsPreviewMode": (121, 2, (11, 0), (), "IsPreviewMode", None),
        "IsPrivateInfoProtected": (124, 2, (11, 0), (), "IsPrivateInfoProtected", None),
        "LastCtrl": (6, 2, (9, 0), (), "LastCtrl", None),
        "PageCount": (7, 2, (3, 0), (), "PageCount", None),
        "ParaShape": (8, 2, (9, 0), (), "ParaShape", None),
        "ParentCtrl": (9, 2, (9, 0), (), "ParentCtrl", None),
        "Path": (10, 2, (8, 0), (), "Path", None),
        "ReadOnlyMode": (14, 2, (11, 0), (), "ReadOnlyMode", None),
        "ScrollPosInfo": (73, 2, (9, 0), (), "ScrollPosInfo", None),
        "SelectionMode": (15, 2, (2, 0), (), "SelectionMode", None),
        "Version": (12, 2, (3, 0), (), "Version", None),
        "ViewProperties": (11, 2, (9, 0), (), "ViewProperties", None),
        "XHwpDocuments": (122, 2, (9, 0), (), "XHwpDocuments", None),
    }
    _prop_map_put_ = {
        "AutoShowHideToolBar" : ((71, LCID, 4, 0),()),
        "CellShape" : ((1, LCID, 4, 0),()),
        "CharShape" : ((2, LCID, 4, 0),()),
        "CurFieldState" : ((13, LCID, 4, 0),()),
        "CurSelectedCtrl" : ((123, LCID, 4, 0),()),
        "EditMode" : ((56, LCID, 4, 0),()),
        "EngineProperties" : ((72, LCID, 4, 0),()),
        "HeadCtrl" : ((3, LCID, 4, 0),()),
        "HyperlinkMode" : ((78, LCID, 4, 0),()),
        "IsEmpty" : ((4, LCID, 4, 0),()),
        "IsModified" : ((5, LCID, 4, 0),()),
        "IsPreviewMode" : ((121, LCID, 4, 0),()),
        "IsPrivateInfoProtected" : ((124, LCID, 4, 0),()),
        "LastCtrl" : ((6, LCID, 4, 0),()),
        "PageCount" : ((7, LCID, 4, 0),()),
        "ParaShape" : ((8, LCID, 4, 0),()),
        "ParentCtrl" : ((9, LCID, 4, 0),()),
        "Path" : ((10, LCID, 4, 0),()),
        "ReadOnlyMode" : ((14, LCID, 4, 0),()),
        "ScrollPosInfo" : ((73, LCID, 4, 0),()),
        "SelectionMode" : ((15, LCID, 4, 0),()),
        "Version" : ((12, LCID, 4, 0),()),
        "ViewProperties" : ((11, LCID, 4, 0),()),
        "XHwpDocuments" : ((122, LCID, 4, 0),()),
    }
    def __iter__(self):
        "Return a Python iterator for this object"
        try:
            ob = self._oleobj_.InvokeTypes(-4,LCID,3,(13, 10),())
        except pythoncom.error:
            raise TypeError("This object does not support enumeration")
        return win32com.client.util.Iterator(ob, None)

class _DHwpCtrlEvents:
    'Event interface for HwpCtrl Control'
    CLSID = CLSID_Sink = IID('{402995CB-FF79-4467-B7D1-4FFA32592B21}')
    coclass_clsid = IID('{BD9C32DE-3155-4691-8972-097D53B10052}')
    _public_methods_ = [] # For COM Server support
    _dispid_to_func_ = {
                1 : "OnNotifyMessage",
                2 : "OnMouseLButtonDown",
                3 : "OnScroll",
        }

    def __init__(self, oobj = None):
        if oobj is None:
            self._olecp = None
        else:
            import win32com.server.util
            from win32com.server.policy import EventHandlerPolicy
            cpc=oobj._oleobj_.QueryInterface(pythoncom.IID_IConnectionPointContainer)
            cp=cpc.FindConnectionPoint(self.CLSID_Sink)
            cookie=cp.Advise(win32com.server.util.wrap(self, usePolicy=EventHandlerPolicy))
            self._olecp,self._olecp_cookie = cp,cookie
    def __del__(self):
        try:
            self.close()
        except pythoncom.com_error:
            pass
    def close(self):
        if self._olecp is not None:
            cp,cookie,self._olecp,self._olecp_cookie = self._olecp,self._olecp_cookie,None,None
            cp.Unadvise(cookie)
    def _query_interface_(self, iid):
        import win32com.server.util
        if iid==self.CLSID_Sink: return win32com.server.util.wrap(self)

    # Event Handlers
    # If you create handlers, they should have the following prototypes:
#    def OnNotifyMessage(self, Msg=defaultNamedNotOptArg, WParam=defaultNamedNotOptArg, LParam=defaultNamedNotOptArg):
#    def OnMouseLButtonDown(self, x=defaultNamedNotOptArg, y=defaultNamedNotOptArg):
#    def OnScroll(self, WParam=defaultNamedNotOptArg, LParam=defaultNamedNotOptArg):


from win32com.client import CoClassBaseClass
class HwpAction(CoClassBaseClass): # A CoClass
    CLSID = IID('{39F4EB62-6155-446D-B210-1F010EED7062}')
    coclass_sources = [
    ]
    coclass_interfaces = [
        DHwpAction,
    ]
    default_interface = DHwpAction

# This CoClass is known by the name 'HWPCONTROL.HwpCtrlCtrl.1'
class HwpCtrl(CoClassBaseClass): # A CoClass
    # HwpCtrl Control
    CLSID = IID('{BD9C32DE-3155-4691-8972-097D53B10052}')
    coclass_sources = [
        _DHwpCtrlEvents,
    ]
    default_source = _DHwpCtrlEvents
    coclass_interfaces = [
        _DHwpCtrl,
    ]
    default_interface = _DHwpCtrl

class HwpCtrlCode(CoClassBaseClass): # A CoClass
    CLSID = IID('{C2452DD4-E559-4BF8-B6BF-ACCC3540A02F}')
    coclass_sources = [
    ]
    coclass_interfaces = [
        DHwpCtrlCode,
    ]
    default_interface = DHwpCtrlCode

class HwpMenu(CoClassBaseClass): # A CoClass
    CLSID = IID('{C9153331-791D-4AF5-AD10-17AC64012A5D}')
    coclass_sources = [
    ]
    coclass_interfaces = [
        DHwpMenu,
    ]
    default_interface = DHwpMenu

class HwpParameterArray(CoClassBaseClass): # A CoClass
    CLSID = IID('{30517C18-C328-4C1E-AE70-4BE894335FB3}')
    coclass_sources = [
    ]
    coclass_interfaces = [
        DHwpParameterArray,
    ]
    default_interface = DHwpParameterArray

class HwpParameterSet(CoClassBaseClass): # A CoClass
    CLSID = IID('{FA18D456-0798-4AD0-AC3A-82A3A7DB1519}')
    coclass_sources = [
    ]
    coclass_interfaces = [
        DHwpParameterSet,
    ]
    default_interface = DHwpParameterSet

RecordMap = {
}

CLSIDToClassMap = {
    '{377C0BC8-E22C-45ED-851A-FBA0208DDC23}' : _DHwpCtrl,
    '{402995CB-FF79-4467-B7D1-4FFA32592B21}' : _DHwpCtrlEvents,
    '{BD9C32DE-3155-4691-8972-097D53B10052}' : HwpCtrl,
    '{5227AA3C-8332-4F0A-9B82-D7D89A5CF395}' : DHwpAction,
    '{39F4EB62-6155-446D-B210-1F010EED7062}' : HwpAction,
    '{E5BBBA9F-732F-450D-8433-1B7AA4F60503}' : DHwpParameterSet,
    '{FA18D456-0798-4AD0-AC3A-82A3A7DB1519}' : HwpParameterSet,
    '{E35C1FFB-BAD5-44B7-A8F2-7B598123C2FC}' : DHwpParameterArray,
    '{30517C18-C328-4C1E-AE70-4BE894335FB3}' : HwpParameterArray,
    '{39E84D56-02DA-4AB0-B0C9-4F8AD94E3DF6}' : DHwpCtrlCode,
    '{C2452DD4-E559-4BF8-B6BF-ACCC3540A02F}' : HwpCtrlCode,
    '{982B2F01-9EAB-4783-AE1D-EA7BAFEFCE2F}' : DHwpMenu,
    '{C9153331-791D-4AF5-AD10-17AC64012A5D}' : HwpMenu,
}
CLSIDToPackageMap = {}
win32com.client.CLSIDToClass.RegisterCLSIDsFromDict( CLSIDToClassMap )
VTablesToPackageMap = {}
VTablesToClassMap = {
}


NamesToIIDMap = {
    '_DHwpCtrl' : '{377C0BC8-E22C-45ED-851A-FBA0208DDC23}',
    '_DHwpCtrlEvents' : '{402995CB-FF79-4467-B7D1-4FFA32592B21}',
    'DHwpAction' : '{5227AA3C-8332-4F0A-9B82-D7D89A5CF395}',
    'DHwpParameterSet' : '{E5BBBA9F-732F-450D-8433-1B7AA4F60503}',
    'DHwpParameterArray' : '{E35C1FFB-BAD5-44B7-A8F2-7B598123C2FC}',
    'DHwpCtrlCode' : '{39E84D56-02DA-4AB0-B0C9-4F8AD94E3DF6}',
    'DHwpMenu' : '{982B2F01-9EAB-4783-AE1D-EA7BAFEFCE2F}',
}


