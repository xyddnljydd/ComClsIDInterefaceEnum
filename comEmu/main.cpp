#include<iostream>
#include<Windows.h>
#include<comdef.h>
#include<rpcdce.h>

#pragma comment(lib, "Rpcrt4.lib")
using namespace std;

void WriteFileBuf(char *buf,SIZE_T size,WCHAR *uid)
{
	char fileName[MAX_PATH] = { 0 };
	sprintf(fileName, "InterfaceFuncs_%ls.txt", uid);
	FILE* f = fopen(fileName, "a+");
	if (f)
	{
		fwrite(buf, size, 1, f);
		fclose(f);
	}
}

void WSExeDemo()
{
	HRESULT hres;
	hres = CoInitializeEx(0, COINIT_MULTITHREADED);

	LPDISPATCH lpDisp;
	CLSID clsidshell;
	hres = CLSIDFromProgID(L"WScript.Shell", &clsidshell);
	if (FAILED(hres))
	{
		printf("CLSIDFromProgID failed %x \n", hres);
		CoUninitialize();
		return;
	}

	hres = CoCreateInstance(clsidshell, NULL, CLSCTX_INPROC_SERVER, IID_IDispatch, (LPVOID*)& lpDisp);
	if (FAILED(hres))
	{
		printf("CoCreateInstance failed %x \n", hres);
		CoUninitialize();
		return;
	}


	LPOLESTR pFuncName = (LPOLESTR)L"Exec";
	DISPID Run;
	hres = lpDisp->GetIDsOfNames(IID_NULL, &pFuncName, 1, LOCALE_SYSTEM_DEFAULT, &Run);
	if (FAILED(hres))
	{
		printf("GetIDsOfNames failed %x \n", hres);
		lpDisp->Release();
		CoUninitialize();
		return;
	}

	VARIANTARG V[1];
	V[0].vt = VT_BSTR;
	V[0].bstrVal = _bstr_t(L"calc.exe");
	DISPPARAMS disParams = { V, NULL, 1, 0 };
	hres = lpDisp->Invoke(Run, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_METHOD, &disParams, NULL, NULL, NULL);
	if (FAILED(hres))
		printf("Invoke failed %x \n", hres);
		
	lpDisp->Release();
	CoUninitialize();
}

std::string GetArgType(UINT type)
{
	std::string typeStr = "";
	switch (type)
	{
	case 0:
		{
			typeStr = "VT_EMPTY";
			break;
		}
	case 1:
		{
			typeStr = "VT_NULL";
			break;
		}
	case 2:
		{
			typeStr = "VT_I2";
			break;
		}
	case 3:
		{
			typeStr = "VT_I4";
			break;
		}
	case 4:
		{
			typeStr = "VT_R4";
			break;
		}
	case 5:
		{
			typeStr = "VT_R8";
			break;
		}
	case 6:
		{
			typeStr = "VT_CY";
			break;
		}
	case 7:
		{
			typeStr = "VT_DATE";
			break;
		}
	case 8:
		{
			typeStr = "VT_BSTR";
			break;
		}
	case 9:
		{
			typeStr = "VT_DISPATCH";
			break;
		}
	case 10:
		{
			typeStr = "VT_ERROR";
			break;
		}
	case 11:
		{
			typeStr = "VT_BOOL";
			break;
		}
	case 12:
		{
			typeStr = "VT_VARIANT";
			break;
		}
	case 13:
		{
			typeStr = "VT_UNKNOWN";
			break;
		}
	case 14:
		{
			typeStr = "VT_DECIMAL";
			break;
		}
	case 16:
		{
			typeStr = "VT_I1";
			break;
		}
	case 17:
		{
			typeStr = "VT_UI1";
			break;
		}
	case 18:
		{
			typeStr = "VT_UI2";
			break;
		}
	case 19:
		{
			typeStr = "VT_UI4";
			break;
		}
	case 20:
		{
			typeStr = "VT_I8";
			break;
		}
	case 21:
		{
			typeStr = "VT_UI8";
			break;
		}
	case 22:
		{
			typeStr = "VT_INT";
			break;
		}
	case 23:
		{
			typeStr = "VT_UINT";
			break;
		}
	case 24:
		{
			typeStr = "VT_VOID";
			break;
		}
	case 25:
		{
			typeStr = "VT_HRESULT";
			break;
		}
	case 26:
		{
			typeStr = "VT_PTR";
			break;
		}
	case 27:
		{
			typeStr = "VT_SAFEARRAY";
			break;
		}
	case 28:
		{
			typeStr = "VT_CARRAY";
			break;
		}
	case 29:
		{
			typeStr = "VT_USERDEFINED";
			break;
		}
	case 30:
		{
		typeStr = "VT_LPSTR";
			break;
		}
	case 31:
		{
			typeStr = "VT_LPWSTR";
			break;
		}
	case 36:
		{
			typeStr = "VT_RECORD";
			break;
		}
	case 37:
		{
			typeStr = "VT_INT_PTR";
			break;
		}
	case 38:
		{
			typeStr = "VT_UINT_PTR";
			break;
		}
	case 64:
		{
			typeStr = "VT_FILETIME";
			break;
		}
	case 65:
		{
			typeStr = "VT_BLOB";
			break;
		}
	case 66:
		{
			typeStr = "VT_STREAM";
			break;
		}
	case 67:
		{
			typeStr = "VT_STORAGE";
			break;
		}
	case 68:
		{
			typeStr = "VT_STREAMED_OBJECT";
			break;
		}
	case 69:
		{
			typeStr = "VT_STORED_OBJECT";
			break;
		}
	case 70:
		{
			typeStr = "VT_BLOB_OBJECT";
			break;
		}
	case 71:
		{
			typeStr = "VT_CF";
			break;
		}
	case 72:
		{
			typeStr = "VT_CLSID";
			break;
		}
	case 73:
		{
			typeStr = "VT_VERSIONED_STREAM";
			break;
		}
	case 0xfff:
		{
			typeStr = "VT_BSTR_BLOB";
			break;
		}
	case 0x1000:
		{
			typeStr = "VT_VECTOR";
			break;
		}
	case 0x2000:
		{
			typeStr = "VT_ARRAY";
			break;
		}
	case 0x4000:
		{
			typeStr = "VT_BYREF";
			break;
		}
	case 0x8000:
		{
			typeStr = "VT_RESERVED";
			break;
		}
	case 0xffff:
		{
			typeStr = "VT_ILLEGAL";
			break;
		}
	default:
		break;
	}
	return typeStr;
}

std::string EmuComInterfaceFuncs(std::wstring progID, std::wstring strClsid, BOOL isLocal)
{
	std::string buf;
	char tmp[MAX_PATH] = { 0 };
	
	CLSID clsidshell;
	LPDISPATCH lpDisp;
	HRESULT hres = E_FAIL;
	hres = CoInitializeEx(0, COINIT_MULTITHREADED);
	if(!progID.empty())
		hres = CLSIDFromProgID(progID.c_str(), &clsidshell);
	else if(!strClsid.empty())
		hres = CLSIDFromString(strClsid.c_str(), &clsidshell);
	if (FAILED(hres))
	{
		memset(tmp, 0, MAX_PATH);
		sprintf(tmp,"CLSIDFromProgID or CLSIDFromString failed %x \n", hres);
		buf += tmp;
		CoUninitialize();
		return buf;
	}

	RPC_CSTR uuidStr = { 0 };
	UuidToStringA(&clsidshell, &uuidStr);
	memset(tmp, 0, MAX_PATH);
	sprintf(tmp,"clsid %s\n", uuidStr);
	buf += tmp;

	//IID demo;
	//UuidFromString((RPC_CSTR)"41904400-be18-11d3-a28b-00104bd35090", &demo);
	if(!isLocal)
		hres = CoCreateInstance(clsidshell, NULL, CLSCTX_INPROC_SERVER, IID_IDispatch, (LPVOID*)& lpDisp);
	else
		hres = CoCreateInstance(clsidshell, NULL, CLSCTX_LOCAL_SERVER, IID_IDispatch, (LPVOID*)& lpDisp);
	if (FAILED(hres))
	{
		memset(tmp, 0, MAX_PATH);
		sprintf(tmp,"CoCreateInstance failed %x \n", hres);
		buf += tmp;
		CoUninitialize();
		return buf;
	}

	ITypeInfo* pTypeInfo = NULL;
	hres = lpDisp->GetTypeInfo(0, LOCALE_SYSTEM_DEFAULT, &pTypeInfo);
	if (FAILED(hres))
	{
		memset(tmp, 0, MAX_PATH);
		sprintf(tmp,"GetTypeInfo failed %x \n", hres);
		buf += tmp;
		lpDisp->Release();
		CoUninitialize();
		return buf;
	}

	ITypeLib* ppTLib;
	UINT pIndex;
	hres = pTypeInfo->GetContainingTypeLib(&ppTLib, &pIndex);
	if (FAILED(hres))
	{
		memset(tmp, 0, MAX_PATH);
		sprintf(tmp,"GetContainingTypeLib failed %x \n", hres);
		buf += tmp;
		lpDisp->Release();
		CoUninitialize();
		return buf;
	}

	memset(tmp, 0, MAX_PATH);
	sprintf(tmp,"total interface %d\n", ppTLib->GetTypeInfoCount());
	buf += tmp;

	for (UINT i = 0; i < ppTLib->GetTypeInfoCount(); i++)
	{
		BSTR bstrName;
		ppTLib->GetDocumentation(i, &bstrName, NULL, NULL, NULL);

		memset(tmp, 0, MAX_PATH);
		if (pIndex == i)
			sprintf(tmp,"\t<==========>current used interfaceName %ls \n", bstrName);
		else
			sprintf(tmp,"\tinterfaceName %ls \n", bstrName);
		buf += tmp;

		ITypeInfo* pTypeInfo2 = NULL;
		ppTLib->GetTypeInfo(i, &pTypeInfo2);


		TYPEATTR* pta;
		hres = pTypeInfo2->GetTypeAttr(&pta);
		if (FAILED(hres))
		{
			memset(tmp, 0, MAX_PATH);
			sprintf(tmp,"GetTypeAttr failed %x \n", hres);
			buf += tmp;
			continue;
		}

		WORD nCount = pta->cFuncs;
		RPC_CSTR currentUuidStr = { 0 };
		UuidToStringA(&pta->guid, &currentUuidStr);
		memset(tmp, 0, MAX_PATH);
		sprintf(tmp,"\tuid %s Funcs Count %d \n", currentUuidStr, nCount);
		buf += tmp;

		if (nCount)
		{
			for (WORD j = 0; j < nCount; j++)
			{
				FUNCDESC* pfd;
				if (SUCCEEDED(pTypeInfo2->GetFuncDesc(j, &pfd)))
				{
					//in fact most functions args are not more than 31
					BSTR	 bstrName[0x20];
					UINT     pcNames = NULL;
					if (SUCCEEDED(pTypeInfo2->GetNames(pfd->memid, bstrName, 0x20, &pcNames)))
					{
						memset(tmp, 0, MAX_PATH);
						sprintf(tmp,"\t\t");
						buf += tmp;
						for (UINT p = 0; p < pcNames; p++)
						{
							memset(tmp, 0, MAX_PATH);
							if (!p)
								sprintf(tmp, "%s %ls(", GetArgType(pfd->elemdescFunc.tdesc.vt).c_str(), bstrName[p]);
							else if (p != pcNames - 1)
								sprintf(tmp, "%s %ls, ", GetArgType(pfd->lprgelemdescParam[p - 1].tdesc.vt).c_str(), bstrName[p]);
							else
								sprintf(tmp, "%s %ls)", GetArgType(pfd->lprgelemdescParam[p - 1].tdesc.vt).c_str(), bstrName[p]);
							buf += tmp;
							if (pcNames == 1)
								buf += ")";//sprintf(tmp, ")");
						}
						memset(tmp, 0, MAX_PATH);
						sprintf(tmp, "\n");
						buf += tmp;
					}
					pTypeInfo2->ReleaseFuncDesc(pfd);
				}
			}
		}
		pTypeInfo2->ReleaseTypeAttr(pta);
	}

	lpDisp->Release();
	CoUninitialize();
	return buf;
}

bool GetIsLocal(WCHAR * achSubKey)
{
	bool isLocal = false;
	WCHAR    achKey[MAX_PATH] = { 0 };   // buffer for subkey name               // size of name string 
	WCHAR    achClass[MAX_PATH] = { 0 };  // buffer for class name 
	DWORD    cchClassName = MAX_PATH;  // size of class string 
	DWORD    cSubKeys = 0;               // number of subkeys 
	DWORD    cbMaxSubKey;              // longest subkey size 
	DWORD    cchMaxClass;              // longest class string 
	DWORD    cValues;              // number of values for key 
	DWORD    cchMaxValue;          // longest value name 
	DWORD    cbMaxValueData;       // longest value data 
	DWORD    cbSecurityDescriptor; // size of security descriptor 
	FILETIME ftLastWriteTime;      // last write time 
	HKEY hKey;
	WCHAR subReg[MAX_PATH] = { L"CLSID\\" };
	wcscat(subReg, achSubKey);

	if (RegOpenKeyExW(HKEY_CLASSES_ROOT, subReg, 0, KEY_READ, &hKey) == ERROR_SUCCESS)
	{
		RegQueryInfoKeyW(hKey, achClass, &cchClassName, NULL, &cSubKeys, &cbMaxSubKey, &cchMaxClass, &cValues, &cchMaxValue, &cbMaxValueData, &cbSecurityDescriptor, &ftLastWriteTime);
		if (cSubKeys)
		{
			for (DWORD i = 0; i < cSubKeys; i++)
			{
				DWORD cbName = MAX_PATH;
				DWORD retCode = RegEnumKeyExW(hKey, i, achKey, &cbName, NULL, NULL, NULL, &ftLastWriteTime);
				if (retCode == ERROR_SUCCESS)
				{
					//printf("\t%ls\n", achKey);
					if (lstrcmpW(achKey,L"LocalServer32") == 0)
					{
						isLocal = true;
						break;
					}
				}
			}
		}
	}
	RegCloseKey(hKey);
	return isLocal;
}


void QueryKey(HKEY hKey)
{
	WCHAR    achKey[MAX_PATH] = { 0 };   // buffer for subkey name               // size of name string 
	WCHAR    achClass[MAX_PATH] = { 0 };  // buffer for class name 
	DWORD    cchClassName = MAX_PATH;  // size of class string 
	DWORD    cSubKeys = 0;               // number of subkeys 
	DWORD    cbMaxSubKey;              // longest subkey size 
	DWORD    cchMaxClass;              // longest class string 
	DWORD    cValues;              // number of values for key 
	DWORD    cchMaxValue;          // longest value name 
	DWORD    cbMaxValueData;       // longest value data 
	DWORD    cbSecurityDescriptor; // size of security descriptor 
	FILETIME ftLastWriteTime;      // last write time 

	RegQueryInfoKeyW(hKey, achClass, &cchClassName, NULL, &cSubKeys, &cbMaxSubKey, &cchMaxClass, &cValues, &cchMaxValue, &cbMaxValueData, &cbSecurityDescriptor, &ftLastWriteTime);
	if (cSubKeys)
	{
		for (DWORD i = 1; i < cSubKeys; i++)
		{
			DWORD cbName = MAX_PATH;
			DWORD retCode = RegEnumKeyExW(hKey, i, achKey, &cbName, NULL, NULL, NULL, &ftLastWriteTime);
			if (retCode == ERROR_SUCCESS)
			{
				bool isLocal = GetIsLocal(achKey);
				printf("%ls isLocal %d\n", achKey, isLocal);
				if (isLocal)
				{
					std::string buf = EmuComInterfaceFuncs(L"", achKey, isLocal);
					WriteFileBuf((char*)buf.c_str(), buf.length(), (WCHAR*)achKey);
				}
			}
		}
	}
}

void EmuRegClsid()
{
	HKEY hKey;
	if (RegOpenKeyEx(HKEY_CLASSES_ROOT, "CLSID", 0, KEY_READ, &hKey) == ERROR_SUCCESS)
	{
		QueryKey(hKey);
	}
	RegCloseKey(hKey);
}


int main()
{
	printf("Waitting for writed ...\n");
	EmuRegClsid();
	printf("Finshed\n");

	//WSExeDemo();
	//std::string buf = "";
	//buf = EmuComInterfaceFuncs(L"",L"{0F87369F-A4E5-4CFC-BD3E-73E6154572DD}",false);
	//buf = EmuComInterfaceFuncs(L"WScript.Shell",L"", false);
	//buf = EmuComInterfaceFuncs(L"Excel.Application",L"", true);
	//WriteFileBuf((char *)buf.c_str(),buf.length(), (WCHAR*)L"");
	system("pause");
	return TRUE;
}