#include <node.h>
#include <v8.h>
#include <windows.h>
#include <string>
#include <vector>
#include <sstream>
#include <iomanip>
#include <propvarutil.h>

#pragma comment(lib, "Propsys.lib")

using namespace v8;
int s_garbageCollectorLevel = 100;
int s_garbageCollectorObjectSize=2000;

Persistent<ObjectTemplate> s_rawTemplateEnum;
Persistent<ObjectTemplate> s_rawTemplate;
Persistent<FunctionTemplate> s_funcTemplate;

class CMyComObject
{
public:
    CMyComObject(IDispatch* obj, bool needToBeReleased)
        : m_needToBeReleased(needToBeReleased),
        m_typeInterface(1),
        m_interface((void*)obj)
    {
    }
    CMyComObject(IEnumVARIANT* obj, bool needToBeReleased)
        : m_needToBeReleased(needToBeReleased),
        m_typeInterface(2),
        m_interface((void*)obj)
    {
    }

    void* m_interface;
    int m_typeInterface;

    bool m_needToBeReleased;

    ~CMyComObject()
    {
        if (m_needToBeReleased && m_interface != NULL)
        {
            if (m_typeInterface == 1)
                ((IDispatch*)m_interface)->Release();
            else if (m_typeInterface == 2)
                ((IEnumVARIANT*)m_interface)->Release();

            m_interface = NULL;
        }
    }
};
std::vector<CMyComObject*> s_comobjects;
std::vector<IDispatch*> dispatches;
Local<Object> WrapMyObject(Isolate* isolate, CMyComObject* myobject);
Local<Value> VariantToValue(Isolate* isolate, VARIANT* pVar);
HRESULT ValueToVariant(Isolate* isolate, Local<Value> value, VARIANT* pVar);

static void SetGarbageCollectorLevel(const v8::FunctionCallbackInfo<v8::Value>& args);
static void DispatchInvoke(const v8::FunctionCallbackInfo<v8::Value>& args);
void DispatchPut(Local<Name> name, Local<Value> value, const PropertyCallbackInfo<void>& info);
static void EnumVariantNext(const v8::FunctionCallbackInfo<v8::Value>& args);
static void EnumVariantIterator(const v8::FunctionCallbackInfo<v8::Value>& args);
static void DispatchToString(const v8::FunctionCallbackInfo<v8::Value>& args);
static void ReleaseInterface(const v8::FunctionCallbackInfo<v8::Value>& args);

void DispatchGet(Local<Name> name, const PropertyCallbackInfo<Value>& info);
void MapGet(Local<Name> name, const PropertyCallbackInfo<Value>& info);
void MapSet(Local<Name> name, Local<Value> value, const PropertyCallbackInfo<Value>& info);
void IndexGet(uint32_t index, const PropertyCallbackInfo<Value>& info);

std::wstring GetErrorMessage(DWORD scode);

std::wstring s_errorMessage;
LONG s_errorCode = 0;
int s_errorLine = 0;
int s_errorCol = 0;
std::string s_errorSourceLine;

Global<Object> s_functions;

	
inline Local<String> v8str(Isolate* isolate, const char* text) {
	Local<String> str;
	if (!text || !String::NewFromUtf8(isolate, text, NewStringType::kNormal).ToLocal(&str)) {
		str = String::Empty(isolate);
	}
	return str;
}

void ReportException(Isolate* isolate, Local<v8::Message> message, Local<v8::Value> exception_obj) {

	HandleScope handle_scope(isolate);
	Local<Context> context = isolate->GetCurrentContext();
	bool enter_context = context.IsEmpty();
	if (enter_context) {
		context = Local<Context>::New(isolate, isolate->GetCurrentContext());
		context->Enter();
	}
	// Converts a V8 value to a C string.
	auto ToCString = [](const v8::String::Utf8Value& value) {
		return *value ? *value : "<string conversion failed>";
		};

	v8::String::Utf8Value exception(isolate, exception_obj);
	const char* exception_string = ToCString(exception);

	int lenwide = MultiByteToWideChar(CP_UTF8, 0, exception_string, -1, NULL, 0);
	if (lenwide > 0)
	{
		wchar_t* ptr = new wchar_t[lenwide + 1];
		memset(ptr, 0, (lenwide + 1) * 2);
		if (ptr)
		{
			MultiByteToWideChar(CP_UTF8, 0, exception_string, -1, ptr, lenwide);
			s_errorMessage = ptr;
			delete[] ptr;
		}
	}


	if (message.IsEmpty()) {
		
	}
	else if (message->GetScriptOrigin().Options().IsWasm())
	{
		int function_index = message->GetWasmFunctionIndex();
		int offset = message->GetStartColumn(context).FromJust();
	}
	else
	{
		s_errorLine = message->GetLineNumber(context).FromMaybe(-1);

		Local<String> sourceline;
		if (message->GetSourceLine(context).ToLocal(&sourceline)) {
			v8::String::Utf8Value sourcelinevalue(isolate, sourceline);
			const char* sourceline_string = ToCString(sourcelinevalue);
			s_errorSourceLine = sourceline_string;
			s_errorCol = message->GetStartColumn(context).FromJust();
		}
	}
	Local<Value> stack_trace_string;
	if (v8::TryCatch::StackTrace(context, exception_obj)
		.ToLocal(&stack_trace_string) &&
		stack_trace_string->IsString()) {
		v8::String::Utf8Value stack_trace(isolate, stack_trace_string.As<String>());
	}
	if (enter_context) context->Exit();
}


CMyComObject* UnwrapMyObject(Local<Object> obj) {
    Local<External> field = obj->GetInternalField(0).As<External>();
    void* ptr = field->Value();
    return static_cast<CMyComObject*>(ptr);
}

std::string ObjectToString(v8::Isolate* isolate, Local<Value> value) {
	String::Utf8Value utf8_value(isolate, value);
	return std::string(*utf8_value);
}

std::wstring ObjectToWString(v8::Isolate* isolate, Local<Value> value)
{
	String::Value val(isolate, value);
	std::wstring wval = (wchar_t*)*val;

	return wval;
}

int ObjectToInt(v8::Isolate* isolate, Local<Value> value)
{
	uint32_t i = 0;
	v8::Maybe<uint32_t> maybe_uint = value->Uint32Value(isolate->GetCurrentContext());
	if (maybe_uint.IsJust()) {
		i = maybe_uint.FromJust();
	}

	return i;
}

std::wstring GetErrorMessage(DWORD scode)
{
	std::wstring desc = L"Could not run method";

	if (HRESULT_FACILITY(scode) == FACILITY_WIN32)
	{
		wchar_t* lpBuffer = NULL;

		DWORD errcode = HRESULT_CODE(scode);

		BOOL bRet = FormatMessageW(FORMAT_MESSAGE_ALLOCATE_BUFFER |
			FORMAT_MESSAGE_FROM_SYSTEM,
			NULL, errcode,
			MAKELANGID(LANG_NEUTRAL,
				SUBLANG_SYS_DEFAULT),
			lpBuffer, 0, NULL);

		if (!bRet)
		{
			wchar_t temp[128];
			wsprintfW(temp, L"Error %X", scode);
			desc = temp;
		}

		if (lpBuffer)
		{
			desc = lpBuffer;
			LocalFree(lpBuffer);
		}
	}
	else if (HRESULT_FACILITY(scode) == FACILITY_CONTROL)
	{
		wchar_t vbscriptDll[MAX_PATH];
		GetSystemDirectoryW(vbscriptDll, MAX_PATH);
		wcscat_s(vbscriptDll, MAX_PATH, L"\\vbscript.dll");
		HMODULE hLib = LoadLibraryExW(vbscriptDll, 0, LOAD_LIBRARY_AS_DATAFILE);

		if (hLib)
		{
			DWORD errcode = HRESULT_CODE(scode);

			wchar_t buf[1024];
			if (LoadStringW(hLib, errcode, buf, 1024))
				desc = buf;

			FreeLibrary(hLib);
		}
	}
	else // Non Win32 Error
	{
		wchar_t temp[128];
		wsprintfW(temp, L"%X", scode);
		desc = temp;
	}

	return desc;
}

static void DispatchEmptyCall(const v8::FunctionCallbackInfo<v8::Value>& args)
{
	args.GetReturnValue().Set(false);
}

bool DispatchFunctionOutArgumentType(LONG id, LONG nArg, IDispatch* pDispatch, VARTYPE* retVarType)
{
	*retVarType = VT_NULL;

	UINT c;
	HRESULT hr1 = pDispatch->GetTypeInfoCount(&c);
	if (c == 0)
		return false;

	ITypeInfo* pTypeInfo;
	HRESULT hr = pDispatch->GetTypeInfo(0, LOCALE_SYSTEM_DEFAULT, &pTypeInfo);
	if (hr == S_OK)
	{
		TYPEATTR* attr;
		hr = pTypeInfo->GetTypeAttr(&attr);

		for (int f = 0; f < attr->cFuncs; f++)
		{
			FUNCDESC* funcdesc;
			HRESULT hr = pTypeInfo->GetFuncDesc(f, &funcdesc);

			if (SUCCEEDED(hr))
			{
				if (funcdesc->memid == id)
				{
					if (nArg >= 0 && nArg < funcdesc->cParams)
					{
						TYPEDESC* td = &funcdesc->lprgelemdescParam[nArg].tdesc;
						if (td->vt == VT_PTR)
						{
							*retVarType = td->lptdesc->vt;
							return true;
						}
					}
				}
			}
		}
	}

	return false;
}

class SetWeakCallbackData
{
public:
	SetWeakCallbackData(void* param, v8::Isolate* isolate, const v8::Local<v8::External>& javascript_object)
		: param(param)
	{
		global.Reset(isolate, javascript_object);
	}
	void* param;
	v8::Global<v8::External> global;
};

void cleanupInterface(const v8::WeakCallbackInfo<SetWeakCallbackData>& data)
{

	SetWeakCallbackData* callback_data = data.GetParameter();

	CMyComObject* pobject = (CMyComObject*)callback_data->param;

	for (unsigned int i = 0; i < s_comobjects.size(); i++)
	{
		if (s_comobjects[i] == pobject)
		{
			CMyComObject* obj = s_comobjects[i];
			s_comobjects.erase(s_comobjects.begin() + i);
			delete obj;

			break;
		}
	}

	callback_data->global.Reset(); // free the V8 reference
	delete callback_data; // delete the heap variable so it isn't leaked
}


Local<Object> WrapMyObject(Isolate* isolate, CMyComObject* myobject)
{
	Local<Object> result;

	bool ret;
	if (myobject->m_typeInterface == 2)
	{
		result = s_rawTemplateEnum.Get(isolate)->NewInstance(isolate->GetCurrentContext()).ToLocalChecked();

		Local<String> str_id = String::NewFromTwoByte(isolate, (const uint16_t*)L"next").ToLocalChecked();

		Local<Object> functions = s_functions.Get(isolate);
		bool exists = functions->HasRealNamedProperty(isolate->GetCurrentContext(), str_id).ToChecked();
		if (!exists)
		{
			functions->Set(isolate->GetCurrentContext(), str_id, FunctionTemplate::New(isolate, EnumVariantNext)->GetFunction(isolate->GetCurrentContext()).ToLocalChecked());
		}

		ret = result->CreateDataProperty(isolate->GetCurrentContext(), str_id, functions->Get(isolate->GetCurrentContext(), str_id).ToLocalChecked()).FromMaybe(false);
	}
	else if (myobject->m_typeInterface == 1)
	{
		IDispatch* pDispatch = (IDispatch*)myobject->m_interface;

		result = s_rawTemplate.Get(isolate)->NewInstance(isolate->GetCurrentContext()).ToLocalChecked();
	}

	Local<External> map_ptr = External::New(isolate, myobject);
	result->SetInternalField(0, map_ptr);

	if (myobject->m_needToBeReleased)
		s_comobjects.push_back(myobject);

	if (s_garbageCollectorLevel > 0 && myobject->m_needToBeReleased)
	{
		int nb = s_comobjects.size();

		if (nb > 300)
		{
			isolate->AdjustAmountOfExternalAllocatedMemory(20000 * 1000 * nb);
		}

		SetWeakCallbackData* callback_data = new SetWeakCallbackData(myobject, isolate, map_ptr);
		callback_data->global.SetWeak<SetWeakCallbackData>(callback_data, cleanupInterface, v8::WeakCallbackType::kParameter);
	}

	return result;
}


static void ReleaseInterface(const v8::FunctionCallbackInfo<v8::Value>& args)
{
	args.GetReturnValue().Set(false);
}

static void EnumVariantIterator(const v8::FunctionCallbackInfo<v8::Value>& args)
{
	Isolate* isolate = args.GetIsolate();

	IDispatch* pDispatch = NULL;

	CMyComObject* obj = UnwrapMyObject(args.Holder());
	if (obj == NULL)
		return;

	if (obj->m_typeInterface == 1)
		pDispatch = (IDispatch*)obj->m_interface;
	else
		return;

	DISPPARAMS dispparms = { NULL, NULL, 0, 0 };
	VARIANT result;
	VariantInit(&result);
	unsigned int nArgErr;
	EXCEPINFO excep;
	memset(&excep, 0, sizeof(excep));

	HRESULT hr = pDispatch->Invoke(DISPID_NEWENUM, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYGET, &dispparms, &result, &excep, &nArgErr);

	if (hr == S_OK)
	{
		if (result.vt == VT_UNKNOWN)
		{
			IEnumVARIANT* pEnumVARIANT = NULL;
			hr = result.punkVal->QueryInterface(IID_IEnumVARIANT, (void**)&pEnumVARIANT);
			if (hr == S_OK)
			{
				Local<Object> theObject = WrapMyObject(args.GetIsolate(), new CMyComObject(pEnumVARIANT, true));

				args.GetReturnValue().Set(theObject);

			}
		}
	}
	VariantClear(&result);
}


static void EnumVariantNext(const v8::FunctionCallbackInfo<v8::Value>& args)
{
	Isolate* isolate = args.GetIsolate();

	IEnumVARIANT* pEnumVARIANT = NULL;

	CMyComObject* obj = UnwrapMyObject(args.Holder());
	if (obj == NULL)
		return;

	if (obj->m_typeInterface == 2)
		pEnumVARIANT = (IEnumVARIANT*)obj->m_interface;
	else
		return;

	VARIANT var;
	VariantInit(&var);
	ULONG celt;

	Local<Object> result = Object::New(isolate);

	bool _done = true;
	if (pEnumVARIANT->Next(1, &var, &celt) == S_OK)
	{
		_done = false;

		Local<String> key = String::NewFromOneByte(isolate, (const uint8_t*)"value").ToLocalChecked();

		if (var.vt == VT_DISPATCH)
		{
			IDispatch* pNewDispatch = NULL;
			HRESULT hr = var.pdispVal->QueryInterface(IID_IDispatch, (void**)&pNewDispatch);
			if (hr == S_OK)
			{
				Local<Object> theObject = WrapMyObject(isolate, new CMyComObject(pNewDispatch, true));
				result->Set(isolate->GetCurrentContext(), key, theObject).FromJust();
			}
		}
		else
			result->Set(isolate->GetCurrentContext(), key, VariantToValue(isolate, &var)).FromJust();

	}
	
	VariantClear(&var);

	Local<String> keyDone = String::NewFromOneByte(isolate, (const uint8_t*)"done").ToLocalChecked();
	Local<Value> done = v8::Boolean::New(isolate, _done);
	result->Set(isolate->GetCurrentContext(), keyDone, done).FromJust();

	args.GetReturnValue().Set(result);

}
static void DispatchToString(const v8::FunctionCallbackInfo<v8::Value>& args)
{
	Isolate* isolate = args.GetIsolate();

	IDispatch* pDispatch = NULL;

	CMyComObject* obj = UnwrapMyObject(args.Holder());
	if (obj == NULL)
		return;

	if (obj->m_typeInterface != 1)
		return;

	pDispatch = (IDispatch*)obj->m_interface;

	DISPID dispid = 0;
	ULONG invokeFlags = DISPATCH_METHOD | DISPATCH_PROPERTYGET;

	int nArgs = (int)args.Length();
	DISPPARAMS dispparms = { NULL, NULL, 0, 0 };

	unsigned int nArgErr;
	EXCEPINFO excep;
	memset(&excep, 0, sizeof(excep));
	VARIANT result;
	VariantInit(&result);

	HRESULT hr = pDispatch->Invoke((DISPID)dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, invokeFlags, &dispparms, &result, &excep, &nArgErr);

	if (hr == S_OK)
	{
		VARIANT dest;
		VariantInit(&dest);
		HRESULT hr = VariantChangeType(&dest, &result, 0, VT_BSTR);
		if (hr == S_OK)
		{
			std::wstring strval = dest.bstrVal;
			Local<Value> ret_str = String::NewFromTwoByte(isolate, (const uint16_t*)LPCWSTR(strval.c_str())).ToLocalChecked();

			args.GetReturnValue().Set(ret_str);
			VariantClear(&result);
			return;
		}

	}
}


static void SetGarbageCollectorLevel(const v8::FunctionCallbackInfo<v8::Value>& args)
{
	Isolate* isolate = args.GetIsolate();

	int nArgs = (int)args.Length();

	if (nArgs >= 1)
	{
		if (args[0]->IsInt32())
		{
			v8::Handle<v8::Integer> _arg = v8::Handle<v8::Integer>::Cast(args[0]);
			s_garbageCollectorLevel = _arg->Value();
		}
	}
	if (nArgs >= 2)
	{
		if (args[1]->IsInt32())
		{
			v8::Handle<v8::Integer> _arg = v8::Handle<v8::Integer>::Cast(args[1]);
			s_garbageCollectorObjectSize = _arg->Value();
		}
	}
}
static void DispatchInvoke(const v8::FunctionCallbackInfo<v8::Value>& args)
{
	Isolate* isolate = args.GetIsolate();

	IDispatch* pDispatch = NULL;
	IEnumVARIANT* pEnumVARIANT = NULL;

	CMyComObject* obj = UnwrapMyObject(args.Holder());
	if (obj == NULL)
		return;

	if (obj->m_typeInterface != 1)
		return;

	pDispatch = (IDispatch*)obj->m_interface;

	Local<Value> str_id = args.Data();

	String::Value id(isolate, str_id);
	std::wstring strId((wchar_t*)*id);

	DISPID dispid = 0;
	ULONG invokeFlags = DISPATCH_METHOD | DISPATCH_PROPERTYGET;

	int pos = strId.rfind('_');
	if (pos > 0)
	{
		std::wstring left = strId.substr(2, pos - 2);
		dispid = std::stoul(left, nullptr, 16);

		std::wstring right = strId.substr(pos + 1);
		invokeFlags = std::stoul(right, nullptr, 16);
	}
	else
		dispid = _wtol(strId.c_str());

	bool returnedDispatch = false;

	int nArgs = (int)args.Length();
	DISPPARAMS dispparms = { NULL, NULL, 0, 0 };

	DISPID dPutID = DISPID_PROPERTYPUT;
	if (invokeFlags == DISPATCH_PROPERTYPUT)
	{
		dispparms.rgdispidNamedArgs = &dPutID;
	}

	VARIANT* outVariants = NULL;
	bool* argIsPtr;

	if (nArgs > 0)
	{
		dispparms.rgvarg = new VARIANT[nArgs];
		dispparms.cArgs = nArgs;

		if (invokeFlags == DISPATCH_PROPERTYPUT)
			dispparms.cNamedArgs = 1;

		outVariants = new VARIANT[nArgs];
		argIsPtr = new bool[nArgs];
	}


	int a = 0;
	for (int i = nArgs - 1; i >= 0; i--)
	{
		VariantInit(&dispparms.rgvarg[i]);
		VariantInit(&outVariants[i]);

		argIsPtr[i] = false;

		if (args[a]->IsObject())
		{
			v8::Handle<v8::Object> theobject = v8::Handle<v8::Object>::Cast(args[a]);

			Local<String> str_value = String::NewFromTwoByte(isolate, (const uint16_t*)L"value").ToLocalChecked();

			bool objectHasValueProperty = theobject->HasRealNamedProperty(isolate->GetCurrentContext(), str_value).ToChecked();

			VARTYPE outType;
			if (objectHasValueProperty && DispatchFunctionOutArgumentType(dispid, a, pDispatch, &outType))
			{
				argIsPtr[i] = true;

				if (outType == VT_VARIANT)
				{
					dispparms.rgvarg[i].vt = VT_BYREF | VT_VARIANT;
					dispparms.rgvarg[i].pvarVal = &outVariants[i];
				}
				else if (outType == VT_BSTR)
				{
					outVariants[i].vt = VT_BSTR;
					outVariants[i].bstrVal = SysAllocString(L"");

					dispparms.rgvarg[i].vt = VT_BYREF | VT_BSTR;
					dispparms.rgvarg[i].pbstrVal = &outVariants[i].bstrVal;
				}
				else if (outType)
				{
					outVariants[i].vt = outType;
					dispparms.rgvarg[i].vt = VT_BYREF | outType;

					if (outType == VT_I2)
						dispparms.rgvarg[i].piVal = &outVariants[i].iVal;
					else if (outType == VT_I4)
						dispparms.rgvarg[i].plVal = &outVariants[i].lVal;
					else if (outType == VT_R4)
						dispparms.rgvarg[i].pfltVal = &outVariants[i].fltVal;
					else if (outType == VT_R8)
						dispparms.rgvarg[i].pdblVal = &outVariants[i].dblVal;
					else if (outType == VT_CY)
						dispparms.rgvarg[i].pcyVal = &outVariants[i].cyVal;
					else if (outType == VT_DATE)
						dispparms.rgvarg[i].pdate = &outVariants[i].date;
					else if (outType == VT_BOOL)
						dispparms.rgvarg[i].pboolVal = &outVariants[i].boolVal;
					else if (outType == VT_DECIMAL)
						dispparms.rgvarg[i].pdecVal = &outVariants[i].decVal;
					else if (outType == VT_I1)
						dispparms.rgvarg[i].pcVal = &outVariants[i].cVal;
					else if (outType == VT_UI1)
						dispparms.rgvarg[i].pbVal = &outVariants[i].bVal;
					else if (outType == VT_UI2)
						dispparms.rgvarg[i].puiVal = &outVariants[i].uiVal;
					else if (outType == VT_UI4)
						dispparms.rgvarg[i].pulVal = &outVariants[i].ulVal;
					else if (outType == VT_INT)
						dispparms.rgvarg[i].pintVal = &outVariants[i].intVal;
					else if (outType == VT_UINT)
						dispparms.rgvarg[i].puintVal = &outVariants[i].uintVal;

				}
			}

		}

		if (!argIsPtr[i])
		{
			if (ValueToVariant(args.GetIsolate(), args[a], &dispparms.rgvarg[i]) != S_OK)
				return;
		}

		a++;
	}


	unsigned int nArgErr;
	EXCEPINFO excep;
	memset(&excep, 0, sizeof(excep));
	VARIANT result;
	VariantInit(&result);

	HRESULT hr = pDispatch->Invoke((DISPID)dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, invokeFlags, &dispparms, &result, &excep, &nArgErr);

	if (hr == S_OK)
	{
		int a = 0;
		for (int i = nArgs - 1; i >= 0; i--)
		{
			if (args[a]->IsObject() && argIsPtr[i])
			{
				v8::Handle<v8::Object> theobject = v8::Handle<v8::Object>::Cast(args[a]);

				Local<String> key = String::NewFromOneByte(isolate, (const uint8_t*)"value").ToLocalChecked();

				theobject->Set(isolate->GetCurrentContext(), key, VariantToValue(isolate, &dispparms.rgvarg[i]));
			}

			a++;
		}

		if (result.vt == VT_DISPATCH || result.vt == VT_UNKNOWN)
		{
			IDispatch* pNewDispatch = NULL;
			if (result.vt == VT_UNKNOWN)
				hr = result.punkVal->QueryInterface(IID_IDispatch, (void**)&pNewDispatch);
			else
				hr = result.pdispVal->QueryInterface(IID_IDispatch, (void**)&pNewDispatch);

			Local<Object> theObject = WrapMyObject(isolate, new CMyComObject(pNewDispatch, true));

			args.GetReturnValue().Set(theObject);
		}
		else
		{

			args.GetReturnValue().Set(VariantToValue(isolate, &result));
		}

		// Clean up the variantargs;
		for (int i = 0; i < nArgs; i++)
		{
			int thisArg = nArgs - i - 1; // reverse order
			VariantClear(dispparms.rgvarg + thisArg);

			VariantClear(outVariants + thisArg);
		}
		if (nArgs > 0)
		{
			delete[] dispparms.rgvarg;
			delete[] outVariants;
			delete[] argIsPtr;
		}

		VariantClear(&result);

		return;
	}

	if (hr == DISP_E_EXCEPTION)
	{
		s_errorMessage = excep.bstrDescription;
		s_errorCode = excep.scode;
	}
	else
	{
		s_errorMessage = GetErrorMessage(hr);
		s_errorCode = hr;
	}

	Local<String> error_message = String::NewFromTwoByte(args.GetIsolate(), (uint16_t*)LPCWSTR(s_errorMessage.c_str()), NewStringType::kInternalized).ToLocalChecked();
	Local<Value> exception = v8::Exception::Error(error_message);

	args.GetIsolate()->ThrowException(exception);
}


HRESULT ValueToVariant(Isolate* isolate, Local<Value> arg, VARIANT* pVar)
{
	VariantInit(pVar);

	if (arg->IsBigInt())
	{
		s_errorMessage = L"Value is BigInt, please convert it to Number.";
		s_errorCode = E_INVALIDARG;

		Local<String> error_message = String::NewFromTwoByte(isolate, (uint16_t*)LPCWSTR(s_errorMessage.c_str()), NewStringType::kInternalized).ToLocalChecked();

		Local<Value> exception = v8::Exception::Error(error_message);

		isolate->ThrowException(exception);

		return S_OK;

	}

	if (arg->IsDate())
	{
		double millisSinceEpoch = (*arg)->ToNumber(isolate->GetCurrentContext()).ToLocalChecked()->Value();
		std::time_t t = static_cast<time_t>(millisSinceEpoch / 1000);

		FILETIME ft;
		LARGE_INTEGER time_value;
		time_value.QuadPart = (t * 10000000LL) + 116444736000000000LL;
		ft.dwLowDateTime = time_value.LowPart;
		ft.dwHighDateTime = time_value.HighPart;

		InitVariantFromFileTime(&ft, pVar);

		return S_OK;
	}

	if (arg->IsBoolean())
	{
		bool v = arg->ToBoolean(isolate)->Value();

		pVar->vt = VT_BOOL;
		pVar->boolVal = v;

		return S_OK;
	}

	if (arg->IsNumber())
	{
		double val = arg->ToNumber(isolate->GetCurrentContext()).ToLocalChecked()->Value();

		pVar->vt = VT_R8;
		pVar->dblVal = val;

		return S_OK;
	}

	if (arg->IsString())
	{
		std::wstring str = ObjectToWString(isolate, arg);

		pVar->vt = VT_BSTR;
		pVar->bstrVal = ::SysAllocString(str.c_str());

		return S_OK;
	}

	if (arg->IsInt32())
	{
		pVar->vt = VT_I4;
		pVar->lVal = arg->ToInt32(isolate->GetCurrentContext()).ToLocalChecked()->Value();

		return S_OK;
	}

	if (arg->IsUint32())
	{
		pVar->vt = VT_UI4;
		pVar->ulVal = arg->ToUint32(isolate->GetCurrentContext()).ToLocalChecked()->Value();

		return S_OK;
	}

	if (arg->IsNull())
	{
		pVar->vt = VT_NULL;
		return S_OK;
	}

	if (arg->IsUndefined())
	{
		pVar->vt = VT_EMPTY;
		return S_OK;
	}

	if (arg->IsUint8Array())
	{
		v8::Handle<v8::Uint8Array> thearray = v8::Handle<v8::Uint8Array>::Cast(arg);

		SAFEARRAYBOUND saBound;
		saBound.lLbound = 0;
		saBound.cElements = thearray->Length();

		pVar->vt = VT_ARRAY | VT_UI1;
		pVar->parray = SafeArrayCreate(VT_UI1, 1, &saBound);

		BYTE* src = NULL;

		HRESULT hr = SafeArrayAccessData(pVar->parray, (void HUGEP**) & src);
		if (SUCCEEDED(hr))
		{
			memcpy(src, thearray->Buffer()->Data(), saBound.cElements);

			SafeArrayUnaccessData(pVar->parray);

			return S_OK;
		}

	}

	if (arg->IsArray())
	{
		v8::Handle<v8::Array> thearray = v8::Handle<v8::Array>::Cast(arg);

		SAFEARRAYBOUND saBound;
		saBound.lLbound = 0;
		saBound.cElements = thearray->Length();

		pVar->vt = VT_ARRAY | VT_VARIANT;
		pVar->parray = SafeArrayCreate(VT_VARIANT, 1, &saBound);

		for (unsigned int i = 0; i < thearray->Length(); i++)
		{
			VARIANT elem;
			VariantInit(&elem);
			ValueToVariant(isolate, thearray->Get(isolate->GetCurrentContext(), i).ToLocalChecked(), &elem);

			LONG Indice = i;
			SafeArrayPutElement(pVar->parray, &Indice, (void*)&elem);
		}
		return S_OK;
	}

	if (arg->IsObject())
	{
		v8::Handle<v8::Object> theobject = v8::Handle<v8::Object>::Cast(arg);
		if (theobject->InternalFieldCount() == 1)
		{
			CMyComObject* obj = UnwrapMyObject(theobject);
			if (obj != NULL && obj->m_typeInterface == 1)
			{
				pVar->vt = VT_DISPATCH;
				IDispatch* disp = (IDispatch*)obj->m_interface;
				disp->QueryInterface(IID_IDispatch, (void**)&pVar->pdispVal);

				return S_OK;
			}
		}
	}

	s_errorMessage=L"Value is invalid";
	s_errorCode = E_INVALIDARG;

	Local<String> error_message = String::NewFromTwoByte(isolate, (uint16_t*)LPCWSTR(s_errorMessage.c_str()), NewStringType::kInternalized).ToLocalChecked();

	Local<Value> exception = v8::Exception::Error(error_message);

	isolate->ThrowException(exception);

	return S_OK;
}
Local<Value> VariantToValue(Isolate* isolate, VARIANT* pVar)
{

	VARTYPE thetype = VT_NULL;

	if (pVar->vt & VT_ARRAY)
		thetype = VT_ARRAY;
	else
		thetype = pVar->vt & VT_TYPEMASK;

	bool byref = false;

	if (pVar->vt & VT_BYREF)
		byref = true;

	if (thetype == VT_VARIANT)
	{
		return VariantToValue(isolate, pVar->pvarVal);
	}
	else if (thetype == VT_NULL)
	{
		Local<Value> result = v8::Null(isolate);
		return result;
	}
	else if (thetype == VT_EMPTY)
	{
		Local<Value> result = v8::Undefined(isolate);
		return result;
	}
	else if (thetype == VT_BSTR)
	{
		std::wstring strval;
		if (byref)
			strval = *pVar->pbstrVal;
		else
			strval = pVar->bstrVal;

		Local<Value> result = String::NewFromTwoByte(isolate, (const uint16_t*)LPCWSTR(strval.c_str())).ToLocalChecked();
		return result;
	}
	else if (thetype == VT_I2 || thetype == VT_I1 || thetype == VT_INT)
	{
		VARIANT dest;
		VariantInit(&dest);
		HRESULT hr = VariantChangeType(&dest, pVar, 0, VT_I4);
		if (hr == S_OK)
		{
			int32_t v = dest.lVal;

			Local<Value> result = Integer::New(isolate, v);
			return result;
		}
	}
	else if (thetype == VT_I4)
	{
		int32_t v;

		if (byref)
			v = *pVar->plVal;
		else
			v = pVar->lVal;

		Local<Value> result = Integer::New(isolate, v);
		return result;

	}
	else if (thetype == VT_UI1 || thetype == VT_UI2 || thetype == VT_UINT)
	{
		VARIANT dest;
		VariantInit(&dest);
		HRESULT hr = VariantChangeType(&dest, pVar, 0, VT_UI4);
		if (hr == S_OK)
		{
			ULONG v = dest.ulVal;

			Local<Value> result = Integer::NewFromUnsigned(isolate, v);
			return result;
		}
	}
	else if (thetype == VT_UI4)
	{
		ULONG v;

		if (byref)
			v = *pVar->pulVal;
		else
			v = pVar->ulVal;

		Local<Value> result = Integer::NewFromUnsigned(isolate, v);
		return result;

	}
	else if (thetype == VT_R8)
	{
		double v;

		if (byref)
			v = *pVar->pdblVal;
		else
			v = pVar->dblVal;

		Local<Value> result = v8::Number::New(isolate, v);
		return result;
	}
	else if (thetype == VT_R4 || thetype == VT_DECIMAL || thetype == VT_CY)
	{
		VARIANT dest;
		VariantInit(&dest);
		HRESULT hr = VariantChangeType(&dest, pVar, 0, VT_R8);
		if (hr == S_OK)
		{
			double v = dest.dblVal;
			Local<Value> result = v8::Number::New(isolate, v);
			return result;
		}
	}
	else if (thetype == VT_BOOL)
	{
		bool v;
		if (byref)
			v = *pVar->pboolVal;
		else
			v = pVar->boolVal;

		Local<Value> result = v8::Boolean::New(isolate, v);
		return result;
	}
	else if (thetype == VT_DATE)
	{
		int ret;
		SYSTEMTIME st;
		if (byref)
			ret = VariantTimeToSystemTime(*pVar->pdate, &st);
		else
			ret = VariantTimeToSystemTime(pVar->date, &st);

		if (ret)
		{
			FILETIME ft;
			if (SystemTimeToFileTime(&st, &ft))
			{
				LARGE_INTEGER ull;
				ull.LowPart = ft.dwLowDateTime;
				ull.HighPart = ft.dwHighDateTime;
				LONGLONG t = (ull.QuadPart / 10000LL) - 11644473600000LL;

				double millisSinceEpoch = (double)t;

				Local<Value> result = v8::Date::New(isolate->GetCurrentContext(), millisSinceEpoch).ToLocalChecked();
				return result;
			}
		}
	}
	else if (thetype == VT_ARRAY)
	{
		if (pVar->parray->cDims == 1)
		{
			int cElements = pVar->parray->rgsabound[0].cElements;
			int lLbound = pVar->parray->rgsabound[0].lLbound;

			VARTYPE vt;
			SafeArrayGetVartype(pVar->parray, &vt);

			if (vt == VT_UI1)
			{
				Local<v8::ArrayBuffer> arrBuf = v8::ArrayBuffer::New(isolate, cElements);
				BYTE* src = NULL;

				HRESULT hr = SafeArrayAccessData(pVar->parray, (void HUGEP**) & src);
				if (SUCCEEDED(hr))
				{
					memcpy(arrBuf->Data(), src, cElements);
					Local<v8::Uint8Array> arr = v8::Uint8Array::New(arrBuf, 0, cElements);

					SafeArrayUnaccessData(pVar->parray);

					return arr;
				}

			}
			else
			{
				Local<v8::Array> arr = v8::Array::New(isolate, cElements);

				UINT elemSize = SafeArrayGetElemsize(pVar->parray);
				BYTE* element = new BYTE[elemSize];

				for (int i = 0; i < cElements; i++)
				{
					LONG Indice = i + lLbound;
					SafeArrayGetElement(pVar->parray, &Indice, (void*)element);

					VARIANT elemVariant;
					VariantInit(&elemVariant);
					elemVariant.vt = vt;

					if (vt == VT_VARIANT)
						elemVariant.pvarVal = (VARIANT*)element;
					else if (vt == VT_I1)
						elemVariant.cVal = *(CHAR*)element;
					else if (vt == VT_I2)
						elemVariant.iVal = *(SHORT*)element;
					else if (vt == VT_I4)
						elemVariant.lVal = *(LONG*)element;
					else if (vt == VT_UI1)
						elemVariant.bVal = *(BYTE*)element;
					else if (vt == VT_UI2)
						elemVariant.uiVal = *(USHORT*)element;
					else if (vt == VT_UI4)
						elemVariant.ulVal = *(ULONG*)element;
					else if (vt == VT_INT)
						elemVariant.intVal = *(INT*)element;
					else if (vt == VT_UINT)
						elemVariant.uintVal = *(UINT*)element;
					else if (vt == VT_BOOL)
						elemVariant.boolVal = *(VARIANT_BOOL*)element;
					else if (vt == VT_DECIMAL)
						memcpy(&elemVariant.decVal, (DECIMAL*)element, sizeof(DECIMAL));
					else if (vt == VT_CY)
						memcpy(&elemVariant.cyVal, (CY*)element, sizeof(CY));
					else if (vt == VT_R4)
						elemVariant.fltVal = *(FLOAT*)element;
					else if (vt == VT_R8)
						elemVariant.dblVal = *(DOUBLE*)element;
					else if (vt == VT_DATE)
						elemVariant.uintVal = *(UINT*)element;
					else if (vt == VT_BSTR)
						elemVariant.bstrVal = *(BSTR*)element;
					else if (vt != VT_EMPTY && vt != VT_NULL)
					{
						Local<String> error_message = String::NewFromOneByte(isolate, (const uint8_t*)"This data type is not supported.").ToLocalChecked();
						Local<Value> exception = v8::Exception::Error(error_message);
						isolate->ThrowException(exception);
					}

					arr->Set(isolate->GetCurrentContext(), i, VariantToValue(isolate, &elemVariant));
				}

				return arr;
			}
		}
	}

	VARIANT vt;
	VariantInit(&vt);
	if(VariantChangeType(&vt,pVar,0,VT_BSTR) ==S_OK)
	{
		std::wstring strval = vt.bstrVal;
		Local<Value> str = String::NewFromTwoByte(isolate, (const uint16_t*)LPCWSTR(strval.c_str())).ToLocalChecked();
		return str;
	}

	Local<String> error_message = String::NewFromOneByte(isolate, (const uint8_t*)"This data type is not supported.").ToLocalChecked();
	Local<Value> exception = v8::Exception::Error(error_message);
	isolate->ThrowException(exception);

	bool v = false;
	Local<Value> res = v8::Boolean::New(isolate, v);
	return res;
}

#define ALL_PARAMETERS_OPTIONALS 1

std::string stringifyTypeDesc(TYPEDESC* typeDesc)
{
	if (typeDesc->vt == VT_PTR)
	{
		return stringifyTypeDesc(typeDesc->lptdesc);
	}
	if (typeDesc->vt == VT_SAFEARRAY)
	{
		return "Array(" + stringifyTypeDesc(typeDesc->lptdesc) + ')';
	}
	if (typeDesc->vt == VT_CARRAY)
	{
		std::string str = stringifyTypeDesc(&typeDesc->lpadesc->tdescElem);
		for (int dim = 0; dim < typeDesc->lpadesc->cDims; dim++)
		{
			char temp[256];
			sprintf(temp,"[%d...%d]", typeDesc->lpadesc->rgbounds[dim].lLbound, typeDesc->lpadesc->rgbounds[dim].cElements + typeDesc->lpadesc->rgbounds[dim].lLbound - 1);
			str += temp;
		}

		return str;
	}

	switch (typeDesc->vt)
	{
	case VT_I2: return "Integer";
	case VT_I4: return "Long";
	case VT_R4: return "Double";
	case VT_R8: return "Double";
	case VT_CY: return "Currency";
	case VT_DATE: return "Date";
	case VT_BSTR: return "String";
	case VT_DISPATCH: return "Object";
	case VT_ERROR: return "Long";
	case VT_BOOL: return "Boolean";
	case VT_VARIANT: return "";
	case VT_UNKNOWN: return "Object";
	case VT_UI1: return "Byte";
	case VT_DECIMAL: return "Double";
	case VT_I1: return "String";
	case VT_UI2: return "Integer";
	case VT_UI4: return "Long";
	case VT_I8: return "Long";
	case VT_UI8: return "Long";
	case VT_INT: return "Integer";
	case VT_UINT: return "Integer";
	case VT_HRESULT: return "Long";
	case VT_VOID: return "Void";
	case VT_LPSTR: return "String";
	case VT_LPWSTR: return "String";
	}

	return "";
}

bool IsDispatchFunction(LONG id, IDispatch* pDispatch, ULONG* flags)
{
	*flags = 0;

	UINT c;
	HRESULT hr1 = pDispatch->GetTypeInfoCount(&c);
	if (c == 0)
		return true;

	ITypeInfo *pTypeInfo;
	HRESULT hr = pDispatch->GetTypeInfo(0, LOCALE_SYSTEM_DEFAULT, &pTypeInfo);

	if (hr == S_OK)
	{
		TYPEATTR* attr;
		hr = pTypeInfo->GetTypeAttr(&attr);

		if (hr == S_OK)
		{
			for (int v = 0; v < attr->cVars; v++)
			{
				VARDESC* vardesc;
				HRESULT hr = pTypeInfo->GetVarDesc(v, &vardesc);

				if (SUCCEEDED(hr))
				{
					if (vardesc->memid == id)
					{
						pTypeInfo->ReleaseTypeAttr(attr);
						pTypeInfo->ReleaseVarDesc(vardesc);
						*flags = DISPATCH_PROPERTYGET;
						return false;
					}
					pTypeInfo->ReleaseVarDesc(vardesc);
				}
			}

			for (int f = 0; f < attr->cFuncs; f++)
			{
				FUNCDESC* funcdesc;
				HRESULT hr = pTypeInfo->GetFuncDesc(f, &funcdesc);

				if (SUCCEEDED(hr))
				{
					if (funcdesc->memid == id)
					{
						*flags = funcdesc->invkind;

						if (funcdesc->invkind == INVOKE_PROPERTYGET)
						{
							int cParams = 0;
							int cParamsOpt = 0;
							for (int c = 1; c <= funcdesc->cParams; c++)
							{
								PARAMDESC* paramdesc = &funcdesc->lprgelemdescParam[c - 1].paramdesc;

								if (paramdesc->wParamFlags & PARAMFLAG_FIN)
								{
									cParams++;

									if (paramdesc->wParamFlags & PARAMFLAG_FOPT)
										cParamsOpt++;
								}
							}

							if (cParams == 0)
							{
								pTypeInfo->ReleaseTypeAttr(attr);
								pTypeInfo->ReleaseFuncDesc(funcdesc);
								return false;
							}
							else if (cParams == cParamsOpt)
							{
								pTypeInfo->ReleaseTypeAttr(attr);
								pTypeInfo->ReleaseFuncDesc(funcdesc);
								return false;
							}

							pTypeInfo->ReleaseTypeAttr(attr);
							pTypeInfo->ReleaseFuncDesc(funcdesc);
							return true;

						}
					}

					pTypeInfo->ReleaseFuncDesc(funcdesc);
				}
			}

			pTypeInfo->ReleaseTypeAttr(attr);
		}
	}

	return true;
}

void IndexGet(uint32_t index, const PropertyCallbackInfo<Value>& info)
{
	Isolate* isolate = info.GetIsolate();

	IDispatch* pDispatch = NULL;

	CMyComObject* obj = UnwrapMyObject(info.Holder());

	if (obj == NULL)
		return;

	if (obj->m_typeInterface == 1)
		pDispatch = (IDispatch*)obj->m_interface;
	else
		return;

	std::wstring desc = L"Could not find property";
	HRESULT hr = S_OK;

	if (pDispatch)
	{
		ITypeInfo* pTypeInfo;
		hr = pDispatch->GetTypeInfo(0, LOCALE_SYSTEM_DEFAULT, &pTypeInfo);

		int cParams = -1;
		int cParamsOpt = -1;
		if (hr == S_OK)
		{
			DISPPARAMS dispparams = { NULL, NULL, 0, 0 };

			VARIANT var;
			dispparams.rgvarg = &var;
			dispparams.cArgs = 1;

			VariantInit(&var);
			var.vt = VT_I4;
			var.lVal = index;

			unsigned int nArgErr;
			EXCEPINFO excep;
			memset(&excep, 0, sizeof(excep));
			VARIANT result;
			VariantInit(&result);

			hr = pDispatch->Invoke((DISPID)0, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYGET, &dispparams, &result, &excep, &nArgErr);

			if (hr == S_OK)
			{
				if (result.vt == VT_DISPATCH)
				{
					IDispatch* pNewDispatch = NULL;
					hr = result.pdispVal->QueryInterface(IID_IDispatch, (void**)&pNewDispatch);

					Local<Object> theObject = WrapMyObject(info.GetIsolate(), new CMyComObject(pNewDispatch, true));

					info.GetReturnValue().Set(theObject);

					VariantClear(&result);

					return;
				}
				else
				{
					info.GetReturnValue().Set(VariantToValue(isolate, &result));
					VariantClear(&result);
					return;
				}
			}

			if (hr == DISP_E_EXCEPTION)
			{
				desc = excep.bstrDescription;
			}
			else
			{
				desc = GetErrorMessage(hr);
			}
		}
	}

	Local<String> error_message = String::NewFromTwoByte(info.GetIsolate(), (uint16_t*)LPCWSTR(desc.c_str()), NewStringType::kInternalized).ToLocalChecked();

	Local<Value> exception = v8::Exception::Error(error_message);
	info.GetIsolate()->ThrowException(exception);

}

void MapGet(Local<Name> name, const PropertyCallbackInfo<Value>& info)
{
	Isolate* isolate = info.GetIsolate();

	if (name->IsSymbol())
	{

		Local<v8::Symbol> mysymbol = Local<v8::Symbol>::Cast(name);

		if (mysymbol->Description(isolate)->IsUndefined())
			return;

		Local<String> str = Local<String>::Cast(mysymbol->Description(isolate));

		std::string symbol_desc = ObjectToString(info.GetIsolate(), str);

		if (symbol_desc.compare("Symbol.iterator") == 0)
		{
			Local<String> str_iterator = String::NewFromTwoByte(isolate, (const uint16_t*)L"@@iterator").ToLocalChecked();

			Local<Object> functions = s_functions.Get(isolate);
			bool exists = functions->HasRealNamedProperty(isolate->GetCurrentContext(), str_iterator).ToChecked();
			if (!exists)
			{
				functions->Set(isolate->GetCurrentContext(), str_iterator, FunctionTemplate::New(isolate, EnumVariantIterator)->GetFunction(isolate->GetCurrentContext()).ToLocalChecked());
			}

			info.GetReturnValue().Set(functions->Get(isolate->GetCurrentContext(), str_iterator).ToLocalChecked());

			return;
		}
		return;
	}

	IDispatch* pDispatch = NULL;

	CMyComObject* obj = UnwrapMyObject(info.Holder());

	if (obj == NULL)
		return;

	if (obj->m_typeInterface == 1)
		pDispatch = (IDispatch*)obj->m_interface;
	else
		return;

	std::string key = ObjectToString(info.GetIsolate(), name.As<String>());
	std::wstring wkey = std::wstring(key.begin(), key.end());
	const wchar_t* pStr = wkey.c_str();
	LPOLESTR szMember = (wchar_t*)wkey.c_str();

	std::wstring desc = L"Could not find property";
	HRESULT hr = S_OK;

	if (pDispatch)
	{
		LONG dispid;
		int cParams = -1;
		int cParamsOpt = -1;
		
		ITypeInfo* pTypeInfo;
		hr = pDispatch->GetTypeInfo(0, LOCALE_SYSTEM_DEFAULT, &pTypeInfo);

		if (hr == S_OK)
		{
			ITypeComp* pTypeComp = NULL;
			hr = pTypeInfo->QueryInterface(IID_ITypeComp, (void**)&pTypeComp);

			if (hr == S_OK)
			{
				SYSKIND sysKind = SYS_WIN32;
#ifdef _WIN64
				sysKind = SYS_WIN64;
#endif
				ULONG hash = LHashValOfNameSys(sysKind, LOCALE_SYSTEM_DEFAULT, szMember);

				ITypeInfo* pTInfo = NULL;
				DESCKIND descKind;
				BINDPTR bindPtr;
				memset(&bindPtr, 0, sizeof(bindPtr));

				hr = pTypeComp->Bind(szMember, hash, INVOKE_FUNC | INVOKE_PROPERTYGET, &pTInfo, &descKind, &bindPtr);
				if (hr == S_OK)
				{
					if (descKind == DESCKIND_FUNCDESC)
					{
						if (bindPtr.lpfuncdesc->invkind == INVOKE_PROPERTYGET)
						{
							dispid = bindPtr.lpfuncdesc->memid;
							cParams = 0;
							cParamsOpt = 0;
							for (int c = 1; c <= bindPtr.lpfuncdesc->cParams; c++)
							{
								PARAMDESC* paramdesc = &bindPtr.lpfuncdesc->lprgelemdescParam[c - 1].paramdesc;

								if (paramdesc->wParamFlags & PARAMFLAG_FIN)
								{
									cParams++;

									if (paramdesc->wParamFlags & PARAMFLAG_FOPT)
										cParamsOpt++;
								}
							}
						}
						
					}
					else if (descKind == DESCKIND_VARDESC)
					{
						dispid = bindPtr.lpvardesc->memid;

						cParams = 0;
						cParamsOpt = 0;

					}

					if (pTInfo != NULL)
					{
						if (descKind == DESCKIND_FUNCDESC && bindPtr.lpfuncdesc != NULL)
							pTInfo->ReleaseFuncDesc(bindPtr.lpfuncdesc);

						if (descKind == DESCKIND_VARDESC && bindPtr.lpvardesc != NULL)
							pTInfo->ReleaseVarDesc(bindPtr.lpvardesc);

						pTInfo->Release();
					}
				}
				pTypeComp->Release();
			}

			pTypeInfo->Release();
		}

		if (cParams == 0)
		{
			DISPPARAMS dispparms = { NULL, NULL, 0, 0 };

			unsigned int nArgErr;

			VARIANT result;
			VariantInit(&result);

			EXCEPINFO excep;
			memset(&excep, 0, sizeof(excep));

			hr = pDispatch->Invoke((DISPID)dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYGET, &dispparms, &result, &excep, &nArgErr);

			if (hr == S_OK)
			{
				if (result.vt == VT_DISPATCH)
				{
					IDispatch* pNewDispatch = NULL;
					hr = result.pdispVal->QueryInterface(IID_IDispatch, (void**)&pNewDispatch);

					Local<Object> theObject = WrapMyObject(info.GetIsolate(), new CMyComObject(pNewDispatch, true));

					info.GetReturnValue().Set(theObject);

					VariantClear(&result);

					return;
				}
				else
				{
					info.GetReturnValue().Set(VariantToValue(isolate, &result));
					VariantClear(&result);
					return;
				}
			}

			if (hr == DISP_E_EXCEPTION)
			{
				s_errorMessage = excep.bstrDescription;
				s_errorCode = excep.scode;
			}
			else
			{
				s_errorMessage = GetErrorMessage(hr);
				s_errorCode = hr;
			}
			desc = s_errorMessage;
		}
		else
		{
			ULONG funcFlags = DISPATCH_METHOD | DISPATCH_PROPERTYGET;

			HRESULT hr = pDispatch->GetIDsOfNames(IID_NULL, &szMember, 1, LOCALE_SYSTEM_DEFAULT, &dispid);

			if (hr == DISP_E_UNKNOWNNAME)
			{
				if (wcsncmp(szMember, L"put", 3) == 0 || wcsncmp(szMember, L"set", 3) == 0)
				{
					LPOLESTR szMemberShort = szMember + 3;
					for (int i = 0; i < 5; i++)
					{
						hr = pDispatch->GetIDsOfNames(IID_NULL, &szMemberShort, 1, LOCALE_SYSTEM_DEFAULT, &dispid);
						funcFlags = DISPATCH_PROPERTYPUT;
						if (SUCCEEDED(hr))
							break;

						if (wcslen(szMemberShort) > 0 && szMemberShort[0] == '_')
						{
							szMemberShort++;
						}
						else
							break;
					}

				}
				else if (wcsncmp(szMember, L"get", 3) == 0)
				{
					LPOLESTR szMemberShort = szMember + 3;
					for (int i = 0; i < 5; i++)
					{
						hr = pDispatch->GetIDsOfNames(IID_NULL, &szMemberShort, 1, LOCALE_SYSTEM_DEFAULT, &dispid);
						funcFlags = DISPATCH_PROPERTYGET;
						if (SUCCEEDED(hr))
							break;

						if (wcslen(szMemberShort) > 0 && szMemberShort[0] == '_')
						{
							szMemberShort++;
						}
						else
							break;
					}
				}
				else
				{
					DISPPARAMS dispparams = { NULL, NULL, 0, 0 };

					VARIANT var;
					dispparams.rgvarg = &var;
					dispparams.cArgs = 1;

					if (ValueToVariant(isolate, name, &var) != S_OK)
					{
						return;
					}

					unsigned int nArgErr;
					EXCEPINFO excep;
					memset(&excep, 0, sizeof(excep));
					VARIANT result;
					VariantInit(&result);

					HRESULT hr = pDispatch->Invoke((DISPID)0, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYGET, &dispparams, &result, &excep, &nArgErr);

					if (hr == S_OK)
					{
						if (result.vt == VT_DISPATCH)
						{
							IDispatch* pNewDispatch = NULL;
							HRESULT hr = result.pdispVal->QueryInterface(IID_IDispatch, (void**)&pNewDispatch);

							Local<Object> theObject = WrapMyObject(info.GetIsolate(), new CMyComObject(pNewDispatch, true));

							info.GetReturnValue().Set(theObject);

							VariantClear(&result);

							return;
						}
						else
						{
							info.GetReturnValue().Set(VariantToValue(isolate, &result));
							VariantClear(&result);
							return;
						}
					}

					if (hr == DISP_E_EXCEPTION)
					{
						s_errorMessage = excep.bstrDescription;
						s_errorCode = excep.scode;
					}
					else
					{
						s_errorMessage = GetErrorMessage(hr);
						s_errorCode = hr;
					}
					desc = s_errorMessage;
				}
			}

			if (SUCCEEDED(hr))
			{
				wchar_t data[128];
				wsprintfW(data,L"f_%x_%x", dispid, funcFlags);
				Local<String> str_id = String::NewFromTwoByte(isolate, (const uint16_t*)LPCWSTR(data)).ToLocalChecked();

				Local<Object> functions = s_functions.Get(isolate);
				bool exists = functions->HasRealNamedProperty(isolate->GetCurrentContext(), str_id).ToChecked();
				if (!exists)
				{
					functions->Set(isolate->GetCurrentContext(), str_id, FunctionTemplate::New(isolate, DispatchInvoke, str_id)->GetFunction(isolate->GetCurrentContext()).ToLocalChecked());
				}

				info.GetReturnValue().Set(functions->Get(isolate->GetCurrentContext(), str_id).ToLocalChecked());

				return;
			}
		}
	}

	Local<String> error_message = String::NewFromTwoByte(info.GetIsolate(), (uint16_t*)LPCWSTR(desc.c_str()), NewStringType::kInternalized).ToLocalChecked();

	Local<Value> exception = v8::Exception::Error(error_message);
	info.GetIsolate()->ThrowException(exception);
}

void DispatchGet(Local<Name> name, const PropertyCallbackInfo<Value>& info)
{
	Isolate* isolate = info.GetIsolate();

	if (name->IsSymbol())
		return;

	IDispatch* pDispatch = NULL;
	IEnumVARIANT* pEnumVARIANT = NULL;

	CMyComObject* obj = UnwrapMyObject(info.Holder());

	if (obj == NULL)
		return;

	if (obj->m_typeInterface == 1)
		pDispatch = (IDispatch*)obj->m_interface;
	else
		return;

	std::string key = ObjectToString(info.GetIsolate(), name.As<String>());
	std::wstring wkey = std::wstring(key.begin(), key.end());
	const wchar_t* pStr = wkey.c_str();
	LPOLESTR szMember = (wchar_t*)wkey.c_str();

	std::wstring desc = L"Could not find property";
	HRESULT hr = S_OK;

	if (pDispatch)
	{
		LONG dispid;

		hr = pDispatch->GetIDsOfNames(IID_NULL, &szMember, 1, LOCALE_SYSTEM_DEFAULT, &dispid);

		EXCEPINFO excep;
		memset(&excep, 0, sizeof(excep));

		if (SUCCEEDED(hr))
		{
			DISPPARAMS dispparms = { NULL, NULL, 0, 0 };

			unsigned int nArgErr;

			VARIANT result;
			VariantInit(&result);

			hr = pDispatch->Invoke((DISPID)dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYGET, &dispparms, &result, &excep, &nArgErr);

			if (hr == S_OK)
			{
				if (result.vt == VT_DISPATCH)
				{
					IDispatch* pNewDispatch = NULL;
					hr = result.pdispVal->QueryInterface(IID_IDispatch, (void**)&pNewDispatch);

					Local<Object> theObject = WrapMyObject(info.GetIsolate(), new CMyComObject(pNewDispatch, true));

					info.GetReturnValue().Set(theObject);

					VariantClear(&result);

					return;
				}
				else
				{
					info.GetReturnValue().Set(VariantToValue(isolate, &result));
					VariantClear(&result);
					return;
				}
			}
		}

		if (hr == DISP_E_EXCEPTION)
		{
			s_errorMessage = excep.bstrDescription;
			s_errorCode = excep.scode;
		}
		else
		{
			s_errorMessage = GetErrorMessage(hr);
			s_errorCode = hr;
		}
		
		desc = s_errorMessage;

		Local<String> error_message = String::NewFromTwoByte(info.GetIsolate(), (uint16_t*)LPCWSTR(desc.c_str()), NewStringType::kInternalized).ToLocalChecked();

		Local<Value> exception = v8::Exception::Error(error_message);
		info.GetIsolate()->ThrowException(exception);
	}
}

void DispatchPut(Local<Name> name, Local<Value> arg, const PropertyCallbackInfo<void>& info)
{
	Isolate* isolate = info.GetIsolate();

	if (name->IsSymbol()) return;

	IDispatch* pDispatch = NULL;

	CMyComObject* obj = UnwrapMyObject(info.Holder());
	if (obj == NULL)
		return;

	if (obj->m_typeInterface == 1)
		pDispatch = (IDispatch*)obj->m_interface;
	else
		return;


	std::string key = ObjectToString(info.GetIsolate(), name.As<String>());
	std::wstring wkey = std::wstring(key.begin(), key.end());
	const wchar_t* pStr = wkey.c_str();
	LPOLESTR szMember = (wchar_t*)wkey.c_str();


	if (pDispatch)
	{

		DISPID dispid;
		HRESULT hr = pDispatch->GetIDsOfNames(IID_NULL, &szMember, 1, LOCALE_SYSTEM_DEFAULT, &dispid);
		if (SUCCEEDED(hr))
		{
			DISPPARAMS dispparams = { NULL, NULL, 0, 0 };

			VARIANT var;

			dispparams.rgvarg = &var;
			dispparams.cArgs = 1;
			dispparams.cNamedArgs = 1;
			DISPID dPutID = DISPID_PROPERTYPUT;
			dispparams.rgdispidNamedArgs = &dPutID;

			if (ValueToVariant(isolate, arg, &var) != S_OK)
			{
				return;
			}

			unsigned int nArgErr;
			EXCEPINFO excep;
			memset(&excep, 0, sizeof(excep));
			VARIANT result;
			VariantInit(&result);


			HRESULT hr = pDispatch->Invoke((DISPID)dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYPUT, &dispparams, &result, &excep, &nArgErr);

			VariantClear(&var);
			VariantClear(&result);

			if (hr != S_OK)
			{
				if (hr == DISP_E_EXCEPTION)
				{
					s_errorMessage = excep.bstrDescription;
					s_errorCode = excep.scode;
				}
				else
				{
					s_errorMessage = GetErrorMessage(hr);
					s_errorCode = hr;
				}

				Local<String> error_message = String::NewFromTwoByte(isolate, (uint16_t*)LPCWSTR(s_errorMessage.c_str()), NewStringType::kInternalized).ToLocalChecked();

				Local<Value> exception = v8::Exception::Error(error_message);

				isolate->ThrowException(exception);
			}
		}
	}


}

void MapSet(Local<Name> name, Local<Value> arg, const PropertyCallbackInfo<Value>& info)
{
	Isolate* isolate = info.GetIsolate();

	if (name->IsSymbol()) return;

	IDispatch* pDispatch = NULL;

	CMyComObject* obj = UnwrapMyObject(info.Holder());
	if (obj == NULL)
		return;

	if (obj->m_typeInterface == 1)
		pDispatch = (IDispatch*)obj->m_interface;
	else
		return;

	std::string key = ObjectToString(info.GetIsolate(), name.As<String>());
	std::wstring wkey = std::wstring(key.begin(), key.end());
	const wchar_t* pStr = wkey.c_str();
	LPOLESTR szMember = (wchar_t*)wkey.c_str();


	if (pDispatch)
	{

		DISPID dispid;
		HRESULT hr = pDispatch->GetIDsOfNames(IID_NULL, &szMember, 1, LOCALE_SYSTEM_DEFAULT, &dispid);
		if (SUCCEEDED(hr))
		{
			DISPPARAMS dispparams = { NULL, NULL, 0, 0 };

			VARIANT var;

			dispparams.rgvarg = &var;
			dispparams.cArgs = 1;
			dispparams.cNamedArgs = 1;
			DISPID dPutID = DISPID_PROPERTYPUT;
			dispparams.rgdispidNamedArgs = &dPutID;

			if (ValueToVariant(isolate, arg, &var) != S_OK)
			{
				return;
			}

			unsigned int nArgErr;
			EXCEPINFO excep;
			memset(&excep, 0, sizeof(excep));
			VARIANT result;
			VariantInit(&result);

			HRESULT hr = pDispatch->Invoke((DISPID)dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYPUT, &dispparams, &result, &excep, &nArgErr);

			VariantClear(&var);
			VariantClear(&result);


			if (hr != S_OK)
			{
				if (hr == DISP_E_EXCEPTION)
				{
					s_errorMessage = excep.bstrDescription;
					s_errorCode = excep.scode;
				}
				else
				{
					s_errorMessage = GetErrorMessage(hr);
					s_errorCode = hr;
				}

				Local<String> error_message = String::NewFromTwoByte(isolate, (uint16_t*)LPCWSTR(s_errorMessage.c_str()), NewStringType::kInternalized).ToLocalChecked();

				Local<Value> exception = v8::Exception::Error(error_message);

				isolate->ThrowException(exception);
			}
		}
	}

	info.GetReturnValue().Set(false);
}

static void CreateObject(const v8::FunctionCallbackInfo<v8::Value>& args) 
{
    v8::Isolate* isolate = args.GetIsolate();

    int nArgs = (int)args.Length();

	if (nArgs == 1 && args[0]->IsString())
	{
		std::wstring str = ObjectToWString(isolate, args[0]);

		CLSID clsid;
		HRESULT hr = CLSIDFromProgID(str.c_str(), &clsid);
		if (hr == S_OK)
		{
			LPUNKNOWN punk = NULL;
			hr = CoCreateInstance(clsid, NULL, CLSCTX_SERVER, IID_IUnknown, (void**)&punk);
			if (hr == S_OK)
			{
				IDispatch* disp;
				hr = punk->QueryInterface(IID_IDispatch, (void**)&disp);

				punk->Release();

				if (hr == S_OK)
				{
					args.GetReturnValue().Set(WrapMyObject(isolate, new CMyComObject(disp, true)));
					return;
				}
			}
		}
	}

    v8::Local<v8::String> error_message = v8::String::NewFromTwoByte(args.GetIsolate(), (uint16_t*)L"incorrect parameter", v8::NewStringType::kInternalized).ToLocalChecked();
    v8::Local<v8::Value> exception = v8::Exception::Error(error_message);
    isolate->ThrowException(exception);
}

extern "C" NODE_MODULE_EXPORT void
NODE_MODULE_INITIALIZER(Local<Object> exports,
	Local<Value> module,
	Local<Context> context)
{    
	Isolate* isolate = context->GetIsolate();
	
	Local<Object> functions = Object::New(isolate);
	s_functions.Reset(isolate,functions);

    v8::Local<v8::Integer> id = v8::Integer::New(isolate, 0);

    HRESULT hr = CoInitialize(NULL);

	Local<FunctionTemplate> clazz = FunctionTemplate::New(isolate, CreateObject);
	clazz->SetClassName(v8str(isolate, "Dispatch"));


	Local<ObjectTemplate>& inst = clazz->InstanceTemplate();
	inst->SetInternalFieldCount(1);
	inst->SetHandler(NamedPropertyHandlerConfiguration(MapGet, MapSet));
	inst->SetHandler(IndexedPropertyHandlerConfiguration(IndexGet));
	inst->SetCallAsFunctionHandler(DispatchInvoke, id);

	s_rawTemplate.Empty();
	s_funcTemplate.Empty();

	s_rawTemplate.Reset(isolate, inst);
	s_funcTemplate.Reset(isolate, clazz);
	
	Local<ObjectTemplate> rawTemplateEnum;
	rawTemplateEnum = ObjectTemplate::New(isolate);
	rawTemplateEnum->SetInternalFieldCount(1);
	rawTemplateEnum->SetCallAsFunctionHandler(DispatchInvoke, id);

	s_rawTemplateEnum.Empty();
	s_rawTemplateEnum.Reset(isolate,rawTemplateEnum);

	exports->Set(context, v8str(isolate, "CreateObject"), clazz->GetFunction(context).ToLocalChecked());

}

