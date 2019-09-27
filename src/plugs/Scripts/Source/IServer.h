#ifndef _ISERVER_H_
#define _ISERVER_H_

typedef struct {
  const GUID *guid;
  ITypeInfo **ppTypeInfo;
} GUIDTYPEINFO;

//Global variables
extern LONG g_nObjs;
extern LONG g_nLocks;
extern IClassFactory MyIClassFactoryObj;
extern const IClassFactoryVtbl MyIClassFactoryVtbl;

//Functions prototypes
HRESULT LoadTypeInfoFromFile(const GUID *guid, ITypeInfo **ppTypeInfo);

HRESULT STDMETHODCALLTYPE Class_QueryInterface(IClassFactory *this, REFIID factoryGuid, void **ppv);
ULONG STDMETHODCALLTYPE Class_AddRef(IClassFactory *this);
ULONG STDMETHODCALLTYPE Class_Release(IClassFactory *this);
HRESULT STDMETHODCALLTYPE Class_CreateInstance(IClassFactory *this, IUnknown *punkOuter, REFIID vTableGuid, void **objHandle);
HRESULT STDMETHODCALLTYPE Class_LockServer(IClassFactory *this, BOOL flock);

BOOL AKD_IsEqualIID(const GUID *rguid1, const GUID *rguid2);

HRESULT WINAPI DllGetClassObject(REFCLSID objGuid, REFIID factoryGuid, void **factoryHandle);
HRESULT WINAPI DllCanUnloadNow(void);
HRESULT WINAPI DllRegisterServer();
HRESULT WINAPI DllUnregisterServer();

void stringFromCLSIDA(char *szString, REFCLSID clsid);
HRESULT IsServerRegisteredA(HINSTANCE hInstance, REFCLSID clsid, const char *pNameTypeLib, const char *pNameProgID, const char *pObjectDescription, char *szDllFile, int nDllFileLen);
HRESULT RegisterServerA(HINSTANCE hInstance, REFCLSID clsid, const char *pNameTypeLib, const char *pNameProgID, const char *pObjectDescription);
HRESULT UnregisterServerA(REFCLSID clsid, REFGUID typelib, const char *pNameProgID);

#endif
