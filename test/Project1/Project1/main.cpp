#include <mshtml.h>    
#include <atlbase.h>    
#include <oleacc.h>   
#include <string>   
#include <comdef.h> 

#include <windows.h>
#include <stdio.h>
#include <stdlib.h>
#include <iostream>

using namespace std;
//
//using std::string;
VOID GetStr(HWND hwndIE)
{
	CoInitialize(NULL);
	CComPtr<IHTMLDocument2> pDoc2;
	string strTemp = "";
	HINSTANCE hinst = LoadLibraryA("OLEACC.DLL");
	if (hinst != NULL) {
		LRESULT lres;
		UINT unMsg = RegisterWindowMessageA("WM_HTML_GETOBJECT");
		SendMessageTimeoutA(hwndIE, unMsg, 0L, 0L, SMTO_ABORTIFHUNG, 1000, (PDWORD_PTR)&lres);
		LPFNOBJECTFROMLRESULT pfObjectFromLresult = (LPFNOBJECTFROMLRESULT)::GetProcAddress(hinst, "ObjectFromLresult");
		if (pfObjectFromLresult != NULL) {
			HRESULT hres;
			hres = (*pfObjectFromLresult)(lres, IID_IHTMLDocument2, 0, (void**)&pDoc2);
			if (SUCCEEDED(hres))
			{
				CComPtr<IHTMLElement> pHtmlElem;
				hres = pDoc2->get_body(&pHtmlElem);
				BSTR bstrText = NULL;
				pHtmlElem->get_innerText(&bstrText);
				_bstr_t _bstrTemp(bstrText, false);
				strTemp = (char*)_bstrTemp;
			}
		}
		FreeLibrary(hinst);
	}
	CoUninitialize();
}



BOOL CALLBACK EnumChildProc(HWND hwnd, LPARAM lParam)
{
	TCHAR buf[100];

	::GetClassName(hwnd, (LPTSTR)&buf, 100);
	if (_tcscmp(buf, _T("Inte.NET Explorer_Server")) == 0)

	{
		*(HWND*)lParam = hwnd;
		return FALSE;
	}
	else
		return TRUE;
};

//You can store the interface pointer in a member variable
//for easier access
void CDlg::OnGetDocInterface(HWND hWnd)
{
	CoInitialize(NULL);

	// Explicitly load MSAA so we know if it's installed
	HINSTANCE hInst = ::LoadLibrary(_T("OLEACC.DLL"));
	if (hInst != NULL)
	{
		if (hWnd != NULL)
		{
			HWND hWndChild = NULL;
			// Get 1st document window
			::EnumChildwindows(hWnd, EnumChildProc, (LPARAM)&hWndChild);
			if (hWndChild)
			{
				CComPtr spDoc;
				LRESULT lRes;

				UINT nMsg = ::RegisterWindowMessage(_T("WM_HTML_GETobject"));
				::SendMessageTimeout(hWndChild, nMsg, 0L, 0L, SMTO_ABORTIFHUNG, 1000, (Dword*)&lRes);

				LPFNOBJECTFROMLRESULT pfObjectFromLresult = (LPFNOBJECTFROMLRESULT)::GetProcAddress(hInst, _T("ObjectFromLresult"));
				if (pfObjectFromLresult != NULL)
				{
					HRESULT hr;
					hr = (*pfObjectFromLresult)(lRes, IID_IHTMLDocument, 0, (void**)&spDoc);
					if (SUCCEEDED(hr))
					{
						CComPtr spDisp;
						CComQIPtr spWin;
						spDoc->get_Script(&spDisp);
						spWin = spDisp;
						spWin->get_document(&spDoc.p);
						// Change background color to red
						spDoc->put_bgColor(CComVariant("red"));
					}
				}
			} // else document not ready
		} // else Internet Explorer is not running
		::FreeLibrary(hInst);
	} // else Active Accessibility is not installed
	CoUninitialize();
}


INT main()
{

	HWND hwndCDM = NULL;
	HWND hwndTemp = NULL;
	HWND hwndIES = NULL;
	try
	{
		while (hwndIES == NULL)
		{
			hwndCDM = FindWindow(TEXT("#32770"), nullptr);
			if (hwndCDM == NULL)
			{
				throw("出错");
			}
			hwndTemp = FindWindowEx(hwndCDM, NULL, nullptr, nullptr);
			hwndTemp = FindWindowEx(hwndTemp, NULL, nullptr, nullptr);
			hwndIES = FindWindowEx(hwndTemp, NULL, TEXT("Internet Explorer_Server"), nullptr);
		}
	}
	catch (string e)
	{
		cout << e << endl;
		return 0;
	}

	GetStr(hwndIES);

	SendMessage(hwndIES, WM_LBUTTONDOWN, MK_LBUTTON, 30 + 34 * 65536);
	SendMessage(hwndIES, WM_LBUTTONDOWN, MK_LBUTTON, 30 + 34 * 65536);
	Sleep(10);
	SendMessage(hwndIES, WM_LBUTTONUP, 0x0000, 30 + 34 * 65536);
	return 0;

}