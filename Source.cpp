#pragma comment(linker,"\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

#include <windows.h>
#include <atlbase.h>
#include <atlwin.h>
#include <wmp.h>

#define NOUI

TCHAR szClassName[] = TEXT("Window");

CComModule _Module;

BEGIN_OBJECT_MAP(ObjectMap)
END_OBJECT_MAP()

BOOL PlayVideo(HWND hWnd, LPCTSTR lpszFilePath)
{
	BOOL bReturn = FALSE;
	CComPtr<IUnknown> pUnknown;
	if (SUCCEEDED(AtlAxGetControl(hWnd, &pUnknown)))
	{
		CComPtr<IWMPPlayer> pIWMPPlayer;
		if ((SUCCEEDED(pUnknown->QueryInterface(__uuidof(IWMPPlayer), (VOID**)&pIWMPPlayer))))
		{
			BSTR bstrText = SysAllocString(lpszFilePath);
			if (SUCCEEDED(pIWMPPlayer->put_URL(bstrText)))
			{
				bReturn = TRUE;
			}
			SysFreeString(bstrText);
			pIWMPPlayer.Release();
		}
		pUnknown.Release();
	}
	return bReturn;
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	static HWND hWindowsMediaPlayerControl;
	switch (msg)
	{
	case WM_CREATE:
		AtlAxWinInit();
		_Module.Init(ObjectMap, ((LPCREATESTRUCT)lParam)->hInstance);
		{
			LPOLESTR lpolestr;
			StringFromCLSID(__uuidof(WindowsMediaPlayer), &lpolestr);
			hWindowsMediaPlayerControl = CreateWindow(TEXT(ATLAXWIN_CLASS), lpolestr, WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hWnd, 0, ((LPCREATESTRUCT)lParam)->hInstance, 0);
			CoTaskMemFree(lpolestr);
		}
#ifdef NOUI
		if (hWindowsMediaPlayerControl)
		{
			CComPtr<IUnknown> pUnknown;
			if (SUCCEEDED(AtlAxGetControl(hWindowsMediaPlayerControl, &pUnknown)))
			{
				CComPtr<IWMPPlayer> pIWMPPlayer;
				if ((SUCCEEDED(pUnknown->QueryInterface(__uuidof(IWMPPlayer), (VOID**)&pIWMPPlayer))))
				{
					BSTR bstrText = SysAllocString(TEXT("none"));
					pIWMPPlayer->put_uiMode(bstrText);
					SysFreeString(bstrText);
					pIWMPPlayer.Release();
				}
				pUnknown.Release();
			}
		}
#endif
		if (hWindowsMediaPlayerControl && ((LPCREATESTRUCT)lParam)->lpCreateParams)
		{
			LPWSTR lpszFilePath = (LPWSTR)((LPCREATESTRUCT)lParam)->lpCreateParams;
			if (PathFileExists(lpszFilePath))
			{
				PlayVideo(hWindowsMediaPlayerControl, lpszFilePath);
			}
		}
		DragAcceptFiles(hWnd, TRUE);
		break;
	case WM_DROPFILES:
		if (DragQueryFile((HDROP)wParam, 0xFFFFFFFF, NULL, 0) == 1)
		{
			TCHAR szFilePath[MAX_PATH];
			DragQueryFile((HDROP)wParam, 0, szFilePath, _countof(szFilePath));
			PlayVideo(hWindowsMediaPlayerControl, szFilePath);
		}
		DragFinish((HDROP)wParam);
		break;
	case WM_SIZE:
		MoveWindow(hWindowsMediaPlayerControl, 0, 0, LOWORD(lParam), HIWORD(lParam), 1);
		break;
	case WM_COMMAND:
		if (LOWORD(wParam) == 1000) // 全画面表示
		{
			if (hWindowsMediaPlayerControl)
			{
				CComPtr<IUnknown> pUnknown;
				if (SUCCEEDED(AtlAxGetControl(hWindowsMediaPlayerControl, &pUnknown)))
				{
					CComPtr<IWMPPlayer> pIWMPPlayer;
					if ((SUCCEEDED(pUnknown->QueryInterface(__uuidof(IWMPPlayer), (VOID**)&pIWMPPlayer))))
					{
						VARIANT_BOOL b;						
						if (SUCCEEDED(pIWMPPlayer->get_fullScreen(&b)))
						{
							if (b == VARIANT_TRUE)
							{
								pIWMPPlayer->put_fullScreen(VARIANT_FALSE);
							}
							else
							{
								pIWMPPlayer->put_fullScreen(VARIANT_TRUE);
							}
						}						
						pIWMPPlayer.Release();
					}
					pUnknown.Release();
				}
			}
		}
		break;	
	case WM_DESTROY:
		DestroyWindow(hWindowsMediaPlayerControl);
		AtlAxWinTerm();
		_Module.Term();
		PostQuitMessage(0);
		break;
	default:
		return DefWindowProc(hWnd, msg, wParam, lParam);
	}
	return 0;
}

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPreInst, LPSTR pCmdLine, int nCmdShow)
{
	MSG msg;
	WNDCLASS wndclass = {
		0,
		WndProc,
		0,
		0,
		hInstance,
		0,
		LoadCursor(0,IDC_ARROW),
		0,
		0,
		szClassName
	};
	RegisterClass(&wndclass);
	int nArgs;
	LPWSTR *szArglist = CommandLineToArgvW(GetCommandLine(), &nArgs);
	HWND hWnd = CreateWindow(
		szClassName,
		TEXT("ドラッグ&ドロップされた動画をメディアプレイヤーを使って再生する"),
		WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN,
		CW_USEDEFAULT,
		0,
		CW_USEDEFAULT,
		0,
		0,
		0,
		hInstance,
		(szArglist && nArgs == 2) ? (LPVOID)szArglist[1] : 0
	);
	LocalFree(szArglist);
	ShowWindow(hWnd, SW_SHOWDEFAULT);
	UpdateWindow(hWnd);
	ACCEL Accel[] = { { FVIRTKEY,VK_F11,1000 } };
	HACCEL hAccel = CreateAcceleratorTable(Accel, 1);
	while (GetMessage(&msg, 0, 0, 0))
	{
		if (!TranslateAccelerator(hWnd, hAccel, &msg))
		{
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}
	}
	DestroyAcceleratorTable(hAccel);
	return (int)msg.wParam;
}