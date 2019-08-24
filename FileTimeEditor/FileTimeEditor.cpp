// FileTimeEditor.cpp : アプリケーションのエントリ ポイントを定義します。
//

#include "stdafx.h"
#include "FileTimeEditor.h"

// グローバル変数:
HINSTANCE	hInst;
LPCTSTR AppName = _T("File Time Editor Ver. 1.0");

// このコード モジュールに含まれる関数の宣言を転送します:
INT_PTR CALLBACK	About(HWND, UINT, WPARAM, LPARAM);

BOOL							GetFileTime_FileName(HWND, LPCTSTR, LPSYSTEMTIME, LPSYSTEMTIME, LPSYSTEMTIME);
BOOL							SetFileTime_FileName(HWND, LPCTSTR, LPSYSTEMTIME, LPSYSTEMTIME, LPSYSTEMTIME);
void							GetDlgDateTime(HWND, int, int, LPSYSTEMTIME);
void							SetDlgDateTime(HWND, LPTSTR, LPSYSTEMTIME, LPSYSTEMTIME, LPSYSTEMTIME);


int APIENTRY _tWinMain(_In_ HINSTANCE hInstance, _In_opt_ HINSTANCE hPrevInstance, _In_ LPTSTR lpCmdLine, _In_ int nCmdShow)
{
	hInst = hInstance;
	return (int)DialogBox(hInst, MAKEINTRESOURCE(IDD_ABOUTBOX), NULL, About);
}

// バージョン情報ボックスのメッセージ ハンドラーです。
INT_PTR CALLBACK About(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
	UNREFERENCED_PARAMETER(lParam);
	switch (message)
	{
	case WM_INITDIALOG:
	{
		HICON hIcon;
		LPTSTR* lplpArgs;
		int Argc;

		hIcon = (HICON)LoadImage(hInst, MAKEINTRESOURCE(IDI_SMALL), IMAGE_ICON, 32, 32, 0);
		SendMessage(hDlg, WM_SETICON, ICON_SMALL, (LPARAM)hIcon);

		lplpArgs = CommandLineToArgvW(GetCommandLine(), &Argc);
		if (Argc == 2) {
			TCHAR strFile[MAX_PATH] = { 0 };
			SYSTEMTIME CreationTime = { 0 }, LastAccessTime = { 0 }, LastWriteTime = { 0 };

			StringCchCopy(strFile, MAX_PATH, lplpArgs[1]);
			GetFileTime_FileName(hDlg, strFile, &CreationTime, &LastAccessTime, &LastWriteTime);
			SetDlgDateTime(hDlg, strFile, &CreationTime, &LastAccessTime, &LastWriteTime);
		}
		return (INT_PTR)FALSE;
	}
	case WM_COMMAND:
		switch (LOWORD(wParam)) {
		case IDC_BTNREF:	// 参照ボタン
		{
			OPENFILENAME ofn = { 0 };	// 「ファイルを開く」ダイアログ用
			TCHAR strFile[MAX_PATH] = { 0 };	// ファイル名受け取り用
			SYSTEMTIME CreationTime = { 0 }, LastAccessTime = { 0 }, LastWriteTime = { 0 };

			ofn.lStructSize = sizeof(OPENFILENAME);
			ofn.hwndOwner = hDlg;
			ofn.lpstrFilter = _T("All files (*.*)\0*.*\0\0");
			ofn.lpstrCustomFilter = NULL;
			ofn.nFilterIndex = 0;
			ofn.lpstrFile = strFile;
			ofn.nMaxFile = MAX_PATH;
			ofn.lpstrTitle = _T("ファイル選択...");
			ofn.Flags = OFN_FILEMUSTEXIST | OFN_HIDEREADONLY;

			// ファイル開いて日時読み込み
			if (GetOpenFileName(&ofn) == 0) {
				return (INT_PTR)FALSE;
			}
			GetFileTime_FileName(hDlg, strFile, &CreationTime, &LastAccessTime, &LastWriteTime);

			// ダイアログ書き換え
			SetDlgDateTime(hDlg, strFile, &CreationTime, &LastAccessTime, &LastWriteTime);
			return (INT_PTR)TRUE;
		}
		case IDCANCEL:	// 終了
			EndDialog(hDlg, LOWORD(wParam));
			return (INT_PTR)TRUE;

		case IDOK:	// 変更
		{
			SYSTEMTIME CreationTime = { 0 }, LastAccessTime = { 0 }, LastWriteTime = { 0 };
			TCHAR strFile[MAX_PATH] = { 0 };

			if (MessageBox(hDlg, _T("ファイルの時刻情報を上書きします\nよろしいですか？"), AppName, MB_OKCANCEL | MB_ICONQUESTION) != IDOK) {
				return (INT_PTR)FALSE;
			}

			// ダイアログの情報を翻訳
			GetDlgDateTime(hDlg, IDC_DATECreate, IDC_TIMECreate, &CreationTime);
			GetDlgDateTime(hDlg, IDC_DATELastAccess, IDC_TIMELastAccess, &LastAccessTime);
			GetDlgDateTime(hDlg, IDC_DATELastWrite, IDC_TIMELastWrite, &LastWriteTime);

			// ファイル名を取得
			GetDlgItemText(hDlg, IDC_EDITFILE, strFile, MAX_PATH);

			// 書き込む
			SetFileTime_FileName(hDlg, strFile, &CreationTime, &LastAccessTime, &LastWriteTime);
			MessageBox(hDlg, _T("完了しました"), AppName, MB_OK | MB_ICONINFORMATION);
			return (INT_PTR)TRUE;
		}
		case IDC_CHECKLocalTime:	// チェックボックス
		{
			TCHAR strFile[MAX_PATH] = { 0 };
			SYSTEMTIME CreationTime = { 0 }, LastAccessTime = { 0 }, LastWriteTime = { 0 };

			GetDlgItemText(hDlg, IDC_EDITFILE, strFile, MAX_PATH);
			GetFileTime_FileName(hDlg, strFile, &CreationTime, &LastAccessTime, &LastWriteTime);
			SetDlgDateTime(hDlg, strFile, &CreationTime, &LastAccessTime, &LastWriteTime);
			return (INT_PTR)FALSE;
		}
		}
		break;
	case WM_DROPFILES:	// ドラッグ & ドロップ
	{
		TCHAR strFile[MAX_PATH] = { 0 };
		SYSTEMTIME CreationTime = { 0 }, LastAccessTime = { 0 }, LastWriteTime = { 0 };
		HDROP hDrop = (HDROP)wParam;
		DragQueryFile(hDrop, 0, strFile, MAX_PATH);
		DragFinish(hDrop);

		GetFileTime_FileName(hDlg, strFile, &CreationTime, &LastAccessTime, &LastWriteTime);
		SetDlgDateTime(hDlg, strFile, &CreationTime, &LastAccessTime, &LastWriteTime);
		return (INT_PTR)FALSE;
	}
	}
	return (INT_PTR)FALSE;
}

// ファイル名からファイル日時を SYSTEMTIME で返す
BOOL GetFileTime_FileName(HWND hDlg, LPCTSTR szFileName, LPSYSTEMTIME Create, LPSYSTEMTIME Access, LPSYSTEMTIME Write)
{
	HANDLE hFile;
	BOOL Result;
	FILETIME FileTimes[3] = { {0} };

	hFile = CreateFile(szFileName, GENERIC_READ, NULL, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
	if (hFile == INVALID_HANDLE_VALUE) {
		TCHAR Message[2048];
		StringCchPrintf(Message, 2048, _T("ファイルを開けません\nReadError: %s"), szFileName);
		MessageBox(hDlg, Message, AppName, MB_ICONERROR);
		return FALSE;
	}
	Result = GetFileTime(hFile, FileTimes + 0, FileTimes + 1, FileTimes + 2);
	CloseHandle(hFile);

	if (Result == 0) {
		TCHAR Message[2048];
		StringCchPrintf(Message, 2048, _T("ファイルを正常に読み込めません\nReadError: %s"), szFileName);
		MessageBox(hDlg, Message, AppName, MB_ICONERROR);
		return FALSE;
	}

	if (IsDlgButtonChecked(hDlg, IDC_CHECKLocalTime) == BST_CHECKED) {
		for (int i = 0; i < 3; i++) {
			FILETIME Temp;
			FileTimeToLocalFileTime(FileTimes + i, &Temp);
			FileTimes[i] = Temp;
		}
	}
	FileTimeToSystemTime(FileTimes + 0, Create);
	FileTimeToSystemTime(FileTimes + 1, Access);
	FileTimeToSystemTime(FileTimes + 2, Write);

	return Result;
}

// ファイル名からファイル時刻設定
BOOL SetFileTime_FileName(HWND hDlg, LPCTSTR szFileName, LPSYSTEMTIME Create, LPSYSTEMTIME Access, LPSYSTEMTIME Write)
{
	HANDLE hFile;
	BOOL Result;
	FILETIME FileTimes[3] = { {0} };

	SystemTimeToFileTime(Create, FileTimes + 0);
	SystemTimeToFileTime(Access, FileTimes + 1);
	SystemTimeToFileTime(Write, FileTimes + 2);

	if (IsDlgButtonChecked(hDlg, IDC_CHECKLocalTime) == BST_CHECKED) {
		for (int i = 0; i < 3; i++) {
			FILETIME Temp;
			LocalFileTimeToFileTime(FileTimes + i, &Temp);
			FileTimes[i] = Temp;
		}
	}

	hFile = CreateFile(szFileName, GENERIC_WRITE, NULL, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
	if (hFile == INVALID_HANDLE_VALUE) {
		TCHAR Message[2048];
		StringCchPrintf(Message, 2048, _T("ファイルを開けません\nWriteError: %s"), szFileName);
		MessageBox(hDlg, Message, AppName, MB_ICONERROR);
		return FALSE;
	}
	Result = SetFileTime(hFile, FileTimes + 0, FileTimes + 1, FileTimes + 2);
	CloseHandle(hFile);

	if (Result == 0) {
		TCHAR Message[2048];
		StringCchPrintf(Message, 2048, _T("ファイルを正常に書き込めません\nWriteError: %s"), szFileName);
		MessageBox(hDlg, Message, AppName, MB_ICONERROR);
		return FALSE;
	}

	return Result;
}

// ダイアログの時刻を読み出す
void GetDlgDateTime(HWND hDlg, int IDDate, int IDTime, LPSYSTEMTIME TargetTime)
{
	SYSTEMTIME Temp;
	DateTime_GetSystemtime(GetDlgItem(hDlg, IDDate), TargetTime);
	DateTime_GetSystemtime(GetDlgItem(hDlg, IDTime), &Temp);
	TargetTime->wHour = Temp.wHour;
	TargetTime->wMinute = Temp.wMinute;
	TargetTime->wSecond = Temp.wSecond;
	return;
}

// ダイアログに時刻を書き込む
void SetDlgDateTime(HWND hDlg, LPTSTR strFile, LPSYSTEMTIME CreationTime, LPSYSTEMTIME LastAccessTime, LPSYSTEMTIME LastWriteTime)
{
	SetDlgItemText(hDlg, IDC_EDITFILE, strFile);
	DateTime_SetSystemtime(GetDlgItem(hDlg, IDC_DATECreate), GDT_VALID, CreationTime);
	DateTime_SetSystemtime(GetDlgItem(hDlg, IDC_TIMECreate), GDT_VALID, CreationTime);
	DateTime_SetSystemtime(GetDlgItem(hDlg, IDC_DATELastAccess), GDT_VALID, LastAccessTime);
	DateTime_SetSystemtime(GetDlgItem(hDlg, IDC_TIMELastAccess), GDT_VALID, LastAccessTime);
	DateTime_SetSystemtime(GetDlgItem(hDlg, IDC_DATELastWrite), GDT_VALID, LastWriteTime);
	DateTime_SetSystemtime(GetDlgItem(hDlg, IDC_TIMELastWrite), GDT_VALID, LastWriteTime);
	return;
}
