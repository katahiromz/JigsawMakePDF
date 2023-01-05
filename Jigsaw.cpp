// ジグソーメイクPDF by katahiromz
// Copyright (C) 2023 片山博文MZ. All Rights Reserved.
// See README.txt and LICENSE.txt.
#include <windows.h>        // Windowsの標準ヘッダ。
#include <windowsx.h>       // Windowsのマクロヘッダ。
#include <commctrl.h>       // 共通コントロールのヘッダ。
#include <commdlg.h>        // 共通ダイアログのヘッダ。
#include <shlobj.h>         // シェルAPIのヘッダ。
#include <shlwapi.h>        // シェル軽量APIのヘッダ。
#include <tchar.h>          // ジェネリックテキストマッピング用のヘッダ。
#include <strsafe.h>        // 安全な文字列操作用のヘッダ (StringC*)
#include <string>           // std::string および std::wstring クラス。
#include <vector>           // std::vector クラス。
#include <map>              // std::map クラス。
#include <stdexcept>        // std::runtime_error クラス。
#include <cassert>          // assertマクロ。
#include <hpdf.h>           // PDF出力用のライブラリlibharuのヘッダ。
#include "TempFile.hpp"     // 一時ファイル操作用のヘッダ。
#include "gpimage.hpp"      // GDI+を用いた画像ファイル入出力ライブラリ。
#include "Susie.hpp"        // Susieプラグインサポート。
#include "color_value.h"    // color_valueライブラリのヘッダ。
#include "MT.h"             // メルセンヌツイスター乱数生成器。
#include "resource.h"       // リソースIDの定義ヘッダ。

// シェアウェア情報。
#ifndef NO_SHAREWARE
    #include "Shareware.hpp"

    SW_Shareware g_shareware(
        /* company registry key */      TEXT("Katayama Hirofumi MZ"),
        /* application registry key */  TEXT("JigsawMakePDF"),
        /* password hash */
        "3f92983c7aff94d3f6d10fccec5205048958ed06f910232a04b124bdaaebe879",
        /* trial days */                10,
        /* salt string */               "jade3YYyR1",
        /* version string */            "0.9.0");
#endif

// 文字列クラス。
#ifdef UNICODE
    typedef std::wstring string_t;
#else
    typedef std::string string_t;
#endif

// シフトJIS コードページ（Shift_JIS）。
#define CP932  932

// わかりやすい項目名を使用する。
enum
{
    IDC_GENERATE = IDOK,
    IDC_EXIT = IDCANCEL,
    IDC_PAGE_SIZE = cmb1,
    IDC_PAGE_DIRECTION = cmb2,
    IDC_FRAME_WIDTH = cmb3,
    IDC_PIECE_SIZE = cmb4,
    IDC_BACKGROUND_IMAGE = edt1,
    IDC_RANDOM_SEED = edt2,
    IDC_RANDOWM_SEED_UPDOWN = scr1,
    IDC_BROWSE = psh1,
    IDC_ERASE_SETTINGS = psh2,
    IDC_LINE_COLOR = edt8,
    IDC_LINE_COLOR_BUTTON = psh8,
};

// Susieプラグイン マネジャー。
SusiePluginManager g_susie;

struct FONT_ENTRY
{
    string_t m_font_name;
    string_t m_pathname;
    int m_index = -1;
};

// ガゾーナラベPDFのメインクラス。
class JigsawMake
{
public:
    HINSTANCE m_hInstance;
    INT m_argc;
    LPTSTR *m_argv;
    std::map<string_t, string_t> m_settings;

    // コンストラクタ。
    JigsawMake(HINSTANCE hInstance, INT argc, LPTSTR *argv);

    // デストラクタ。
    ~JigsawMake()
    {
    }

    // データをリセットする。
    void Reset();
    // ダイアログを初期化する。
    void InitDialog(HWND hwnd);
    // ダイアログからデータへ。
    BOOL DataFromDialog(HWND hwnd);
    // データからダイアログへ。
    BOOL DialogFromData(HWND hwnd);
    // レジストリからデータへ。
    BOOL DataFromReg(HWND hwnd);
    // データからレジストリへ。
    BOOL RegFromData(HWND hwnd);
    // タグを置き換える。
    bool SubstituteTags(HWND hwnd, string_t& str, const string_t& pathname,
                        INT iImage, INT cImages, INT iPage, INT cPages, bool is_output);

    // メインディッシュ処理。
    string_t JustDoIt(HWND hwnd);
};

// グローバル変数。
HINSTANCE g_hInstance = NULL; // インスタンス。
TCHAR g_szAppName[256] = TEXT(""); // アプリ名。
HICON g_hIcon = NULL; // アイコン（大）。
HICON g_hIconSm = NULL; // アイコン（小）。

// リソース文字列を読み込む。
LPTSTR doLoadString(INT nID)
{
    static TCHAR s_szText[1024];
    s_szText[0] = 0;
    LoadString(NULL, nID, s_szText, _countof(s_szText));
    return s_szText;
}

// 文字列の前後の空白を削除する。
void str_trim(LPWSTR text)
{
    StrTrimW(text, L" \t\r\n\x3000");
}

// ローカルのファイルのパス名を取得する。
LPCTSTR findLocalFile(LPCTSTR filename)
{
    // 現在のプログラムのパスファイル名を取得する。
    static TCHAR szPath[MAX_PATH];
    GetModuleFileName(NULL, szPath, _countof(szPath));

    // ファイルタイトルをfilenameで置き換える。
    PathRemoveFileSpec(szPath);
    PathAppend(szPath, filename);
    if (PathFileExists(szPath))
        return szPath;

    // 一つ上のフォルダへ。
    PathRemoveFileSpec(szPath);
    PathRemoveFileSpec(szPath);
    PathAppend(szPath, filename);
    if (PathFileExists(szPath))
        return szPath;

    // さらに一つ上のフォルダへ。
    PathRemoveFileSpec(szPath);
    PathRemoveFileSpec(szPath);
    PathRemoveFileSpec(szPath);
    PathAppend(szPath, filename);
    if (PathFileExists(szPath))
        return szPath;

    return NULL; // 見つからなかった。
}

// 不正な文字列が入力された。
void OnInvalidString(HWND hwnd, INT nItemID, INT nFieldId, INT nReasonId)
{
    SetFocus(GetDlgItem(hwnd, nItemID));
    string_t field = doLoadString(nFieldId);
    string_t reason = doLoadString(nReasonId);
    TCHAR szText[256];
    StringCchPrintf(szText, _countof(szText), doLoadString(IDS_INVALIDSTRING), field.c_str(), reason.c_str());
    MessageBox(hwnd, szText, g_szAppName, MB_ICONERROR);
}

// コンボボックスのテキストを取得する。
BOOL getComboText(HWND hwnd, INT id, LPTSTR text, INT cchMax)
{
    text[0] = 0;

    HWND hCombo = GetDlgItem(hwnd, id);
    INT iSel = ComboBox_GetCurSel(hCombo);
    if (iSel == CB_ERR) // コンボボックスに選択項目がなければ
    {
        // そのままテキストを取得する。
        ComboBox_GetText(hCombo, text, cchMax);
    }
    else
    {
        // リストからテキストを取得する。長さのチェックあり。
        if (ComboBox_GetLBTextLen(hCombo, iSel) >= cchMax)
        {
            StringCchCopy(text, cchMax, doLoadString(IDS_TEXTTOOLONG));
            return FALSE;
        }
        else
        {
            ComboBox_GetLBText(hCombo, iSel, text);
        }
    }

    return TRUE;
}

// コンボボックスのテキストを設定する。
BOOL setComboText(HWND hwnd, INT id, LPCTSTR text)
{
    // テキストに一致する項目を取得する。
    HWND hCombo = GetDlgItem(hwnd, id);
    INT iItem = ComboBox_FindStringExact(hCombo, -1, text);
    if (iItem == CB_ERR) // 一致する項目がなければ
        ComboBox_SetText(hCombo, text); // そのままテキストを設定。
    else
        ComboBox_SetCurSel(hCombo, iItem); // 一致する項目を選択。
    return TRUE;
}

// ワイド文字列をANSI文字列に変換する。
LPSTR ansi_from_wide(UINT codepage, LPCWSTR wide)
{
    static CHAR s_ansi[1024];

    // コードページで表示できない文字はゲタ文字（〓）にする。
    static const char utf8_geta[] = "\xE3\x80\x93";
    static const char cp932_geta[] = "\x81\xAC";
    const char *geta = NULL;
    if (codepage == CP_ACP || codepage == CP932)
    {
        geta = cp932_geta;
    }
    else if (codepage == CP_UTF8)
    {
        geta = utf8_geta;
    }

    WideCharToMultiByte(codepage, 0, wide, -1, s_ansi, _countof(s_ansi), geta, NULL);
    return s_ansi;
}

// ANSI文字列をワイド文字列に変換する。
LPWSTR wide_from_ansi(UINT codepage, LPCSTR ansi)
{
    static WCHAR s_wide[1024];
    MultiByteToWideChar(codepage, 0, ansi, -1, s_wide, _countof(s_wide));
    return s_wide;
}

// mm単位からピクセル単位への変換。
double pixels_from_mm(double mm, double dpi = 72)
{
    return dpi * mm / 25.4;
}

// ピクセル単位からmm単位への変換。
double mm_from_pixels(double pixels, double dpi = 72)
{
    return 25.4 * pixels / dpi;
}

// ファイル名に使えない文字を下線文字に置き換える。
void validate_filename(string_t& filename)
{
    for (auto& ch : filename)
    {
        if (wcschr(L"\\/:*?\"<>|", ch) != NULL)
            ch = L'_';
    }
}

// 有効な画像ファイルかを確認する。
bool isValidImageFile(LPCTSTR filename)
{
    // 有効なファイルか？
    if (!PathFileExists(filename) || PathIsDirectory(filename))
        return false;

    // GDI+などで読み込める画像ファイルか？
    if (gpimage_is_valid_extension(filename))
        return true;

    // Susie プラグインで読み込める画像ファイルか？
    if (g_susie.is_loaded())
    {
        auto ansi = ansi_from_wide(CP_ACP, filename);
        auto dotext = PathFindExtensionA(ansi);
        if (g_susie.is_dotext_supported(dotext))
            return true;
    }

    return false; // 読み込めない。
}

// 画像を読み込む。
HBITMAP
doLoadPic(LPCWSTR filename, int* width = NULL, int* height = NULL, float* dpi = NULL)
{
    // DPI値を初期化する。
    if (dpi)
        *dpi = 0;

    // GDI+などで読み込みを試みる。
    HBITMAP hbm = gpimage_load(filename, width, height, dpi);
    if (hbm)
        return hbm;

    // Susieプラグインを試す。SusieではANSI文字列を使用。
    auto ansi = ansi_from_wide(CP_ACP, filename);
    hbm = g_susie.load_image(ansi);
    if (hbm == NULL)
    {
        // ANSI文字列では表現できないパスファイル名かもしれない。
        // 一時ファイルを使用して再度挑戦。
        LPCWSTR dotext = PathFindExtension(filename);
        TempFile temp_file;
        temp_file.init(L"GN2", dotext);
        if (CopyFile(filename, temp_file.make(), FALSE))
        {
            ansi = ansi_from_wide(CP_ACP, temp_file.get());
            hbm = g_susie.load_image(ansi);
        }
    }
    if (hbm)
    {
        // 必要な情報を取得する。
        BITMAP bm;
        GetObject(hbm, sizeof(bm), &bm);
        if (width)
            *width = bm.bmWidth;
        if (height)
            *height = bm.bmHeight;
        return hbm;
    }

    return NULL; // 失敗。
}

// コンストラクタ。
JigsawMake::JigsawMake(HINSTANCE hInstance, INT argc, LPTSTR *argv)
    : m_hInstance(hInstance)
    , m_argc(argc)
    , m_argv(argv)
{
    // データをリセットする。
    Reset();
}

// 既定値。
#define IDC_PAGE_SIZE_DEFAULT doLoadString(IDS_A4)
#define IDC_PAGE_DIRECTION_DEFAULT doLoadString(IDS_LANDSCAPE)
#define IDC_FRAME_WIDTH_DEFAULT TEXT("10")
#define IDC_PIECE_SIZE_DEFAULT TEXT("3.0")
#define IDC_RANDOM_SEED_DEFAULT TEXT("0")
#define IDC_BACKGROUND_IMAGE_DEFAULT TEXT("")
#define IDC_LINE_COLOR_DEFAULT TEXT("#FF0080")

// データをリセットする。
void JigsawMake::Reset()
{
#define SETTING(id) m_settings[TEXT(#id)]
    SETTING(IDC_PAGE_SIZE) = IDC_PAGE_SIZE_DEFAULT;
    SETTING(IDC_PAGE_DIRECTION) = IDC_PAGE_DIRECTION_DEFAULT;
    SETTING(IDC_FRAME_WIDTH) = IDC_FRAME_WIDTH_DEFAULT;
    SETTING(IDC_PIECE_SIZE) = IDC_PIECE_SIZE_DEFAULT;
    SETTING(IDC_RANDOM_SEED) = IDC_RANDOM_SEED_DEFAULT;
    SETTING(IDC_BACKGROUND_IMAGE) = IDC_BACKGROUND_IMAGE_DEFAULT;
    SETTING(IDC_LINE_COLOR) = IDC_LINE_COLOR_DEFAULT;
}

// ダイアログを初期化する。
void JigsawMake::InitDialog(HWND hwnd)
{
    // IDC_PAGE_SIZE: 用紙サイズ。
    SendDlgItemMessage(hwnd, IDC_PAGE_SIZE, CB_ADDSTRING, 0, (LPARAM)doLoadString(IDS_A3));
    SendDlgItemMessage(hwnd, IDC_PAGE_SIZE, CB_ADDSTRING, 0, (LPARAM)doLoadString(IDS_A4));
    SendDlgItemMessage(hwnd, IDC_PAGE_SIZE, CB_ADDSTRING, 0, (LPARAM)doLoadString(IDS_A5));
    SendDlgItemMessage(hwnd, IDC_PAGE_SIZE, CB_ADDSTRING, 0, (LPARAM)doLoadString(IDS_B4));
    SendDlgItemMessage(hwnd, IDC_PAGE_SIZE, CB_ADDSTRING, 0, (LPARAM)doLoadString(IDS_B5));

    // IDC_PAGE_DIRECTION: ページの向き。
    SendDlgItemMessage(hwnd, IDC_PAGE_DIRECTION, CB_ADDSTRING, 0, (LPARAM)doLoadString(IDS_PORTRAIT));
    SendDlgItemMessage(hwnd, IDC_PAGE_DIRECTION, CB_ADDSTRING, 0, (LPARAM)doLoadString(IDS_LANDSCAPE));

    // IDC_FRAME_WIDTH: 外枠(mm)。
    SendDlgItemMessage(hwnd, IDC_FRAME_WIDTH, CB_ADDSTRING, 0, (LPARAM)TEXT("0"));
    SendDlgItemMessage(hwnd, IDC_FRAME_WIDTH, CB_ADDSTRING, 0, (LPARAM)TEXT("10"));
    SendDlgItemMessage(hwnd, IDC_FRAME_WIDTH, CB_ADDSTRING, 0, (LPARAM)TEXT("15"));
    SendDlgItemMessage(hwnd, IDC_FRAME_WIDTH, CB_ADDSTRING, 0, (LPARAM)TEXT("20"));

    // IDC_PIECE_SIZE: ピースのサイズ(cm)。
    SendDlgItemMessage(hwnd, IDC_PIECE_SIZE, CB_ADDSTRING, 0, (LPARAM)TEXT("1.5"));
    SendDlgItemMessage(hwnd, IDC_PIECE_SIZE, CB_ADDSTRING, 0, (LPARAM)TEXT("2.0"));
    SendDlgItemMessage(hwnd, IDC_PIECE_SIZE, CB_ADDSTRING, 0, (LPARAM)TEXT("3.0"));
    SendDlgItemMessage(hwnd, IDC_PIECE_SIZE, CB_ADDSTRING, 0, (LPARAM)TEXT("4.0"));

    // IDC_RANDOWM_SEED_UPDOWN: 乱数の種のスピンコントロール。
    SendDlgItemMessage(hwnd, IDC_RANDOWM_SEED_UPDOWN, UDM_SETRANGE, 0, MAKELPARAM(0x7FFF, 0));
}

// ダイアログからデータへ。
BOOL JigsawMake::DataFromDialog(HWND hwnd)
{
    TCHAR szText[MAX_PATH];

    // コンボボックスからデータを取得する。
#define GET_COMBO_DATA(id) do { \
    getComboText(hwnd, (id), szText, _countof(szText)); \
    str_trim(szText); \
    m_settings[TEXT(#id)] = szText; \
} while (0)
    GET_COMBO_DATA(IDC_PAGE_SIZE);
    GET_COMBO_DATA(IDC_PAGE_DIRECTION);
    GET_COMBO_DATA(IDC_FRAME_WIDTH);
    GET_COMBO_DATA(IDC_PIECE_SIZE);
#undef GET_COMBO_DATA

    auto piece_size = SETTING(IDC_PIECE_SIZE);
    if (piece_size.empty())
    {
        SETTING(IDC_PIECE_SIZE) = IDC_PIECE_SIZE_DEFAULT;
        OnInvalidString(hwnd, IDC_PIECE_SIZE, IDS_FIELD_PIECE_SIZE, IDS_REASON_EMPTY_STRING);
        return FALSE;
    }
    wchar_t *endptr;
    double value = wcstod(piece_size.c_str(), &endptr);
    if (*endptr != 0 || value <= 0)
    {
        SETTING(IDC_PIECE_SIZE) = IDC_PIECE_SIZE_DEFAULT;
        OnInvalidString(hwnd, IDC_PIECE_SIZE, IDS_FIELD_PIECE_SIZE, IDS_REASON_POSITIVE_REAL);
        return FALSE;
    }

    ::GetDlgItemText(hwnd, IDC_RANDOM_SEED, szText, _countof(szText));
    str_trim(szText);
    SETTING(IDC_RANDOM_SEED) = szText;

    ::GetDlgItemText(hwnd, IDC_BACKGROUND_IMAGE, szText, _countof(szText));
    str_trim(szText);
    if (!isValidImageFile(szText))
    {
        SETTING(IDC_BACKGROUND_IMAGE) = IDC_BACKGROUND_IMAGE_DEFAULT;
        OnInvalidString(hwnd, IDC_BACKGROUND_IMAGE, IDS_BACKGROUND, IDS_REASON_VALID_IMAGE);
        return FALSE;
    }
    SETTING(IDC_BACKGROUND_IMAGE) = szText;

    ::GetDlgItemText(hwnd, IDC_LINE_COLOR, szText, _countof(szText));
    str_trim(szText);
    if (szText[0] == 0)
    {
        m_settings[TEXT("IDC_LINE_COLOR")] = IDC_LINE_COLOR_DEFAULT;
        ::SetFocus(::GetDlgItem(hwnd, IDC_LINE_COLOR));
        OnInvalidString(hwnd, IDC_LINE_COLOR, IDS_FIELD_TEXT_COLOR, IDS_REASON_EMPTY_STRING);
        return FALSE;
    }
    auto ansi = ansi_from_wide(CP_ACP, szText);
    auto color_value = color_value_parse(ansi);
    if (color_value == -1)
    {
        m_settings[TEXT("IDC_LINE_COLOR")] = IDC_LINE_COLOR_DEFAULT;
        ::SetFocus(::GetDlgItem(hwnd, IDC_LINE_COLOR));
        OnInvalidString(hwnd, IDC_LINE_COLOR, IDS_FIELD_TEXT_COLOR, IDS_REASON_VALID_COLOR);
        return FALSE;
    }
    m_settings[TEXT("IDC_LINE_COLOR")] = szText;

    // チェックボックスからデータを取得する。
#define GET_CHECK_DATA(id) do { \
    if (IsDlgButtonChecked(hwnd, id) == BST_CHECKED) \
        m_settings[TEXT(#id)] = doLoadString(IDS_YES); \
    else \
        m_settings[TEXT(#id)] = doLoadString(IDS_NO); \
} while (0)
#undef GET_CHECK_DATA

    return TRUE;
}

// データからダイアログへ。
BOOL JigsawMake::DialogFromData(HWND hwnd)
{
    // コンボボックスへデータを設定する。
#define SET_COMBO_DATA(id) \
    setComboText(hwnd, (id), m_settings[TEXT(#id)].c_str());
    SET_COMBO_DATA(IDC_PAGE_SIZE);
    SET_COMBO_DATA(IDC_PAGE_DIRECTION);
    SET_COMBO_DATA(IDC_FRAME_WIDTH);
    SET_COMBO_DATA(IDC_PIECE_SIZE);
#undef SET_COMBO_DATA

    SetDlgItemText(hwnd, IDC_RANDOM_SEED, SETTING(IDC_RANDOM_SEED).c_str());
    SetDlgItemText(hwnd, IDC_BACKGROUND_IMAGE, SETTING(IDC_BACKGROUND_IMAGE).c_str());
    SetDlgItemText(hwnd, IDC_LINE_COLOR, SETTING(IDC_LINE_COLOR).c_str());

    // チェックボックスへデータを設定する。
#define SET_CHECK_DATA(id) do { \
    if (m_settings[TEXT(#id)] == doLoadString(IDS_YES)) \
        CheckDlgButton(hwnd, (id), BST_CHECKED); \
    else \
        CheckDlgButton(hwnd, (id), BST_UNCHECKED); \
} while (0)
#undef SET_CHECK_DATA

    return TRUE;
}

// レジストリからデータへ。
BOOL JigsawMake::DataFromReg(HWND hwnd)
{
    // ソフト固有のレジストリキーを開く。
    HKEY hKey;
    RegOpenKeyEx(HKEY_CURRENT_USER, TEXT("Software\\Katayama Hirofumi MZ\\JigsawMakePDF"), 0, KEY_READ, &hKey);
    if (!hKey)
        return FALSE; // 開けなかった。

    // レジストリからデータを取得する。
    TCHAR szText[MAX_PATH];
#define GET_REG_DATA(id) do { \
    szText[0] = 0; \
    DWORD cbText = sizeof(szText); \
    LONG error = RegQueryValueEx(hKey, TEXT(#id), NULL, NULL, (LPBYTE)szText, &cbText); \
    if (error == ERROR_SUCCESS) { \
        SETTING(id) = szText; \
    } \
} while(0)
    GET_REG_DATA(IDC_PAGE_SIZE);
    GET_REG_DATA(IDC_PAGE_DIRECTION);
    GET_REG_DATA(IDC_FRAME_WIDTH);
    GET_REG_DATA(IDC_PIECE_SIZE);
    GET_REG_DATA(IDC_RANDOM_SEED);
    GET_REG_DATA(IDC_BACKGROUND_IMAGE);
    GET_REG_DATA(IDC_LINE_COLOR);
#undef GET_REG_DATA

    // レジストリキーを閉じる。
    RegCloseKey(hKey);
    return TRUE; // 成功。
}

// データからレジストリへ。
BOOL JigsawMake::RegFromData(HWND hwnd)
{
    HKEY hCompanyKey = NULL, hAppKey = NULL;

    // 会社固有のレジストリキーを作成または開く。
    RegCreateKey(HKEY_CURRENT_USER, TEXT("Software\\Katayama Hirofumi MZ"), &hCompanyKey);
    if (hCompanyKey == NULL)
        return FALSE; // 失敗。

    // ソフト固有のレジストリキーを作成または開く。
    RegCreateKey(hCompanyKey, TEXT("JigsawMakePDF"), &hAppKey);
    if (hAppKey == NULL)
    {
        RegCloseKey(hCompanyKey);
        return FALSE; // 失敗。
    }

    // レジストリにデータを設定する。
#define SET_REG_DATA(id) do { \
    auto& str = m_settings[TEXT(#id)]; \
    DWORD cbText = (str.size() + 1) * sizeof(WCHAR); \
    RegSetValueEx(hAppKey, TEXT(#id), 0, REG_SZ, (LPBYTE)str.c_str(), cbText); \
} while(0)
    SET_REG_DATA(IDC_PAGE_SIZE);
    SET_REG_DATA(IDC_PAGE_DIRECTION);
    SET_REG_DATA(IDC_FRAME_WIDTH);
    SET_REG_DATA(IDC_PIECE_SIZE);
    SET_REG_DATA(IDC_RANDOM_SEED);
    SET_REG_DATA(IDC_BACKGROUND_IMAGE);
    SET_REG_DATA(IDC_LINE_COLOR);
#undef SET_REG_DATA

    // レジストリキーを閉じる。
    RegCloseKey(hAppKey);
    RegCloseKey(hCompanyKey);

    return TRUE; // 成功。
}

// 文字列中に見つかった部分文字列をすべて置き換える。
template <typename T_STR>
inline bool
str_replace(T_STR& str, const T_STR& from, const T_STR& to)
{
    bool ret = false;
    size_t i = 0;
    for (;;) {
        i = str.find(from, i);
        if (i == T_STR::npos)
            break;
        ret = true;
        str.replace(i, from.size(), to);
        i += to.size();
    }
    return ret;
}
template <typename T_STR>
inline bool
str_replace(T_STR& str,
            const typename T_STR::value_type *from,
            const typename T_STR::value_type *to)
{
    return str_replace(str, T_STR(from), T_STR(to));
}

// ファイルサイズを取得する。
DWORD get_file_size(const string_t& filename)
{
    WIN32_FIND_DATA find;
    HANDLE hFind = FindFirstFile(filename.c_str(), &find);
    if (hFind == INVALID_HANDLE_VALUE)
        return 0; // エラー。
    FindClose(hFind);
    if (find.nFileSizeHigh)
        return 0; // 大きすぎるのでエラー。
    return find.nFileSizeLow; // ファイルサイズ。
}

// libHaruのエラーハンドラの実装。
void hpdf_error_handler(HPDF_STATUS error_no, HPDF_STATUS detail_no, void* user_data) {
    char message[1024];
    StringCchPrintfA(message, _countof(message), "error: error_no = %04X, detail_no = %d",
                     UINT(error_no), INT(detail_no));
    throw std::runtime_error(message);
}

// 長方形を描画する。
void hpdf_draw_box(HPDF_Page page, double x, double y, double width, double height)
{
    HPDF_Page_MoveTo(page, x, y);
    HPDF_Page_LineTo(page, x, y + height);
    HPDF_Page_LineTo(page, x + width, y + height);
    HPDF_Page_LineTo(page, x + width, y);
    HPDF_Page_ClosePath(page);
    HPDF_Page_Stroke(page);
}

// 画像を描く。
bool hpdf_draw_image(HPDF_Doc pdf, HPDF_Page page, double x, double y, double width, double height, const string_t& filename, bool keep_aspect)
{
    // ファイルサイズを取得する。
    DWORD file_size = get_file_size(filename);
    if (file_size == 0)
        return false;

    // 画像をHBITMAPとして読み込む。
    int image_width, image_height;
    HBITMAP hbm = doLoadPic(filename.c_str(), &image_width, &image_height);
    if (hbm == NULL)
        return false;

    // JPEG/TIFFではないビットマップ画像については特別扱いで画像の品質を向上させることができる。
    bool png_is_better = false;
    auto dotext = PathFindExtension(filename.c_str());
    if (lstrcmpi(dotext, TEXT(".png")) == 0 ||
        lstrcmpi(dotext, TEXT(".gif")) == 0 ||
        lstrcmpi(dotext, TEXT(".bmp")) == 0 ||
        lstrcmpi(dotext, TEXT(".dib")) == 0 ||
        lstrcmpi(dotext, TEXT(".ico")) == 0)
    {
        png_is_better = true;
    }

    TempFile tempfile;
    for (;;)
    {
        file_size = 0;

        // 一時ファイルを作成し、画像を保存する。
        if (png_is_better)
            tempfile.init(TEXT("GN2"), TEXT(".png"));
        else
            tempfile.init(TEXT("GN2"), TEXT(".jpg"));
        if (!gpimage_save(tempfile.make(), hbm))
            break;

        // ファイルサイズを取得する。
        file_size = get_file_size(tempfile.get());
        break;
    }

    // ファイルサイズがゼロなら、失敗。
    if (file_size == 0)
    {
        ::DeleteObject(hbm);
        return false; // 失敗。
    }

    // 画像を読み込む。
    HPDF_Image image;
    if (png_is_better)
        image = HPDF_LoadPngImageFromFile(pdf, ansi_from_wide(CP_ACP, tempfile.get()));
    else
        image = HPDF_LoadJpegImageFromFile(pdf, ansi_from_wide(CP_ACP, tempfile.get()));
    if (image == NULL)
    {
        ::DeleteObject(hbm);
        return false; // 失敗。
    }

    if (keep_aspect)
    {
        // 画像サイズとアスペクト比に従って処理を行う。
        double aspect1 = (double)image_width / (double)image_height;
        double aspect2 = width / height;
        double stretch_width, stretch_height;
        // アスペクト比に従ってセルいっぱいに縮小する。
        if (aspect1 <= aspect2)
        {
            stretch_height = height;
            stretch_width = height * aspect1;
        }
        else
        {
            stretch_width = width;
            stretch_height = width / aspect1;
        }
        // ゼロにならないように補正する。
        if (stretch_width <= 0)
            stretch_width = 1;
        if (stretch_height <= 0)
            stretch_height = 1;
        // 画像を配置する。
        double dx = (width - stretch_width) / 2;
        double dy = (height - stretch_height) / 2;
        HPDF_Page_DrawImage(page, image, x + dx, y + dy, stretch_width, stretch_height);
    }
    else
    {
        // 画像を配置する。
        HPDF_Page_DrawImage(page, image, x, y, width, height);
    }

    // HBITMAPを破棄する。
    ::DeleteObject(hbm);

    return true;
}

// テキストを描画する。
void hpdf_draw_text(HPDF_Page page, HPDF_Font font, double font_size,
                    const char *text,
                    double x, double y, double width, double height,
                    int draw_box = 0)
{
    // フォントサイズを制限。
    if (font_size > HPDF_MAX_FONTSIZE)
        font_size = HPDF_MAX_FONTSIZE;

    // 長方形を描画する。
    if (draw_box == 1)
    {
        hpdf_draw_box(page, x, y, width, height);
    }

    // 長方形に収まるフォントサイズを計算する。
    double text_width, text_height;
    for (;;)
    {
        // フォントとフォントサイズを指定。
        HPDF_Page_SetFontAndSize(page, font, font_size);

        // テキストの幅と高さを取得する。
        text_width = HPDF_Page_TextWidth(page, text);
        text_height = HPDF_Page_GetCurrentFontSize(page);

        // テキストが長方形に収まるか？
        if (text_width <= width && text_height <= height)
        {
            // x,yを中央そろえ。
            x += (width - text_width) / 2;
            y += (height - text_height) / 2;
            break;
        }

        // フォントサイズを少し小さくして再計算。
        font_size *= 0.8;
    }

    // テキストを描画する。
    HPDF_Page_BeginText(page);
    {
        // ベースラインからdescentだけずらす。
        double descent = -HPDF_Font_GetDescent(font) * font_size / 1000.0;
        HPDF_Page_TextOut(page, x, y + descent, text);
    }
    HPDF_Page_EndText(page);

    // 長方形を描画する。
    if (draw_box == 2)
    {
        hpdf_draw_box(page, x, y, text_width, text_height);
    }
}

// 曲線を描画する。
void draw_curve(HPDF_Page page, double px0, double py0, double px1, double py1, double ratio, double slant = 0.0)
{
    double dx = px1 - px0;
    double dy = py1 - py0;
    double mid_x = (px0 + px1) / 2 - ratio * dy;
    double mid_y = (py0 + py1) / 2 + ratio * dx;
    mid_x += slant * dx;
    mid_y += slant * dy;
    HPDF_Page_CurveTo(page, px0, py0, mid_x, mid_y, px1, py1);
}

struct HCellParams
{
    double tab_size;
    double qx1, qx2;
    double slant1, slant2;
    double delta;

    HCellParams(int seed, double px0, double px1, double size)
    {
        switch (seed % 8)
        {
        case 0:
            tab_size = size * 0.25;
            qx1 = (2 * px0 + 1 * px1) / 3;
            qx2 = (1 * px0 + 2 * px1) / 3;
            slant1 = 0.6;
            slant2 = 0;
            delta = 0.8;
            break;
        case 1:
            tab_size = size * 0.25;
            qx1 = (3 * px0 + 2 * px1) / 5;
            qx2 = (2 * px0 + 3 * px1) / 5;
            slant1 = 0.7;
            slant2 = 0.2;
            delta = 0.8;
            break;
        case 2:
            tab_size = size * 0.2;
            qx1 = (5 * px0 + 3 * px1) / 8;
            qx2 = (3 * px0 + 5 * px1) / 8;
            slant1 = 0.6;
            slant2 = 0;
            delta = 0.8;
            break;
        case 3:
            tab_size = size * 0.2;
            qx1 = (3 * px0 + 2 * px1) / 5;
            qx2 = (2 * px0 + 3 * px1) / 5;
            slant1 = 0.75;
            slant2 = 0;
            delta = 0.8;
            break;
        case 4:
            tab_size = size * 0.3;
            qx1 = (3.5 * px0 + 2 * px1) / 5.5;
            qx2 = (2 * px0 + 3.5 * px1) / 5.5;
            slant1 = 0.75;
            slant2 = 0.2;
            delta = 0.4;
            break;
        case 5:
            tab_size = size * 0.2;
            qx1 = (3.5 * px0 + 2 * px1) / 5.5;
            qx2 = (2 * px0 + 3.5 * px1) / 5.5;
            slant1 = 0.75;
            slant2 = 0;
            delta = 0.4;
            break;
        case 6:
            tab_size = size * 0.27;
            qx1 = (4 * px0 + 2 * px1) / 6;
            qx2 = (2 * px0 + 4 * px1) / 6;
            slant1 = 0.75;
            slant2 = -0.1;
            delta = 0.7;
            break;
        case 7:
            tab_size = size * 0.3;
            qx1 = (3.5 * px0 + 2 * px1) / 5.5;
            qx2 = (2 * px0 + 3.5 * px1) / 5.5;
            slant1 = 0.75;
            slant2 = 0;
            delta = 0.4;
            break;
        }
    }
};

struct VCellParams
{
    double tab_size;
    double qy1, qy2;
    double slant1, slant2;
    double delta;

    VCellParams(int seed, double py0, double py1, double size)
    {
        switch (seed % 8)
        {
        case 0:
            tab_size = size * 0.25;
            qy1 = (2 * py0 + 1 * py1) / 3;
            qy2 = (1 * py0 + 2 * py1) / 3;
            slant1 = 0.6;
            slant2 = 0;
            delta = 0.8;
            break;
        case 1:
            tab_size = size * 0.25;
            qy1 = (3 * py0 + 2 * py1) / 5;
            qy2 = (2 * py0 + 3 * py1) / 5;
            slant1 = 0.7;
            slant2 = 0.2;
            delta = 0.8;
            break;
        case 2:
            tab_size = size * 0.2;
            qy1 = (5 * py0 + 3 * py1) / 8;
            qy2 = (3 * py0 + 5 * py1) / 8;
            slant1 = 0.6;
            slant2 = 0;
            delta = 0.8;
            break;
        case 3:
            tab_size = size * 0.2;
            qy1 = (3 * py0 + 2 * py1) / 5;
            qy2 = (2 * py0 + 3 * py1) / 5;
            slant1 = 0.75;
            slant2 = 0;
            delta = 0.8;
            break;
        case 4:
            tab_size = size * 0.3;
            qy1 = (3.5 * py0 + 2 * py1) / 5.5;
            qy2 = (2 * py0 + 3.5 * py1) / 5.5;
            slant1 = 0.75;
            slant2 = 0.2;
            delta = 0.4;
            break;
        case 5:
            tab_size = size * 0.2;
            qy1 = (3.5 * py0 + 2 * py1) / 5.5;
            qy2 = (2 * py0 + 3.5 * py1) / 5.5;
            slant1 = 0.75;
            slant2 = 0;
            delta = 0.4;
            break;
        case 6:
            tab_size = size * 0.27;
            qy1 = (4 * py0 + 2 * py1) / 6;
            qy2 = (2 * py0 + 4 * py1) / 6;
            slant1 = 0.75;
            slant2 = -0.1;
            delta = 0.7;
            break;
        case 7:
            tab_size = size * 0.3;
            qy1 = (3.5 * py0 + 2 * py1) / 5.5;
            qy2 = (2 * py0 + 3.5 * py1) / 5.5;
            slant1 = 0.75;
            slant2 = 0;
            delta = 0.4;
            break;
        }
    }
};

// カット線を描画する（通常ピース）。
void hpdf_draw_cut_lines_normal(HPDF_Doc pdf, HPDF_Page page, double x, double y, double width, double height, int rows, int columns, int seed)
{
    // メルセンヌツイスター乱数生成器を初期化。
    init_genrand(seed);

    // ピースのサイズ。
    double cell_width = width / columns;
    double cell_height = height / rows;

    // 横線を描く。
    for (INT iRow = 1; iRow < rows; ++iRow)
    {
        int sign = (iRow & 1) ? 1 : -1;
        for (INT iColumn = 0; iColumn < columns; ++iColumn)
        {
            double px0 = x + (iColumn + 0) * cell_width;
            double px1 = x + (iColumn + 1) * cell_width;
            double py = y + iRow * cell_height;
            double px_mid = (px0 + px1) / 2;
            HCellParams params(genrand_int31(), px0, px1, cell_height);
            HPDF_Page_MoveTo(page, px0, py);
            draw_curve(page, px0, py, params.qx1, py, -params.delta * sign, params.slant1);
            draw_curve(page, params.qx1, py, px_mid, py + params.tab_size * sign, 0.8 * sign, params.slant2);
            draw_curve(page, px_mid, py + params.tab_size * sign, params.qx2, py, 0.8 * sign, -params.slant2);
            draw_curve(page, params.qx2, py, px1, py, -params.delta * sign, -params.slant1);
            HPDF_Page_Stroke(page);
            sign = -sign;
        }
    }

    // 縦線を描く。
    for (INT iColumn = 1; iColumn < columns; ++iColumn)
    {
        int sign = (iColumn & 1) ? 1 : -1;
        for (INT iRow = 0; iRow < rows; ++iRow)
        {
            double py0 = y + (iRow + 0) * cell_height;
            double py1 = y + (iRow + 1) * cell_height;
            double px = x + iColumn * cell_width;
            double py_mid = (py0 + py1) / 2;
            VCellParams params(genrand_int31(), py0, py1, cell_width);
            HPDF_Page_MoveTo(page, px, py0);
            draw_curve(page, px, py0, px, params.qy1, -params.delta * sign, params.slant1);
            draw_curve(page, px, params.qy1, px - params.tab_size * sign, py_mid, 0.8 * sign, -params.slant2);
            draw_curve(page, px - params.tab_size * sign, py_mid, px, params.qy2, 0.8 * sign, params.slant2);
            draw_curve(page, px, params.qy2, px, py1, -params.delta * sign, -params.slant1);
            HPDF_Page_Stroke(page);
            sign = -sign;
        }
    }
}

// カット線を描画する。
void hpdf_draw_cut_lines(HPDF_Doc pdf, HPDF_Page page, double x, double y, double width, double height, int rows, int columns, int seed)
{
    // 外枠を描画する。
    if (x != 0 || y != 0)
    {
        hpdf_draw_box(page, x, y, width, height);
    }

    hpdf_draw_cut_lines_normal(pdf, page, x, y, width, height, rows, columns, seed);
}

BOOL Aspect_OnInitDialog(HWND hwnd, HWND hwndFocus, LPARAM lParam)
{
    CheckDlgButton(hwnd, rad1, BST_CHECKED);
    return TRUE;
}

void Aspect_OnCommand(HWND hwnd, int id, HWND hwndCtl, UINT codeNotify)
{
    switch (id)
    {
    case IDOK:
        if (IsDlgButtonChecked(hwnd, rad1) == BST_CHECKED)
            EndDialog(hwnd, rad1);
        else if (IsDlgButtonChecked(hwnd, rad2) == BST_CHECKED)
            EndDialog(hwnd, rad2);
        else
            EndDialog(hwnd, IDOK);
        break;
    case IDCANCEL:
        EndDialog(hwnd, id);
        break;
    }
}

// アスペクト比のダイアログプロシージャ。
INT_PTR CALLBACK AspectDlgProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    switch (uMsg)
    {
        HANDLE_MSG(hwnd, WM_INITDIALOG, Aspect_OnInitDialog);
        HANDLE_MSG(hwnd, WM_COMMAND, Aspect_OnCommand);
    }
    return 0;
}

// メインディッシュ処理。
string_t JigsawMake::JustDoIt(HWND hwnd)
{
    string_t ret;
    // PDFオブジェクトを作成する。
    HPDF_Doc pdf = HPDF_New(hpdf_error_handler, NULL);
    if (!pdf)
        return L"";

    try
    {
        // エンコーディング 90ms-RKSJ-H, 90ms-RKSJ-V, 90msp-RKSJ-H, EUC-H, EUC-V が利用可能となる
        HPDF_UseJPEncodings(pdf);

        // 日本語フォントの MS-(P)Mincyo, MS-(P)Gothic が利用可能となる
        HPDF_UseJPFonts(pdf);

        // 用紙の向き。
        HPDF_PageDirection direction;
        if (SETTING(IDC_PAGE_DIRECTION) == doLoadString(IDS_PORTRAIT))
            direction = HPDF_PAGE_PORTRAIT;
        else if (SETTING(IDC_PAGE_DIRECTION) == doLoadString(IDS_LANDSCAPE))
            direction = HPDF_PAGE_LANDSCAPE;
        else
            direction = HPDF_PAGE_LANDSCAPE;

        // ページサイズ。
        HPDF_PageSizes page_size;
        if (SETTING(IDC_PAGE_SIZE) == doLoadString(IDS_A3))
            page_size = HPDF_PAGE_SIZE_A3;
        else if (SETTING(IDC_PAGE_SIZE) == doLoadString(IDS_A4))
            page_size = HPDF_PAGE_SIZE_A4;
        else if (SETTING(IDC_PAGE_SIZE) == doLoadString(IDS_A5))
            page_size = HPDF_PAGE_SIZE_A5;
        else if (SETTING(IDC_PAGE_SIZE) == doLoadString(IDS_B4))
            page_size = HPDF_PAGE_SIZE_B4;
        else if (SETTING(IDC_PAGE_SIZE) == doLoadString(IDS_B5))
            page_size = HPDF_PAGE_SIZE_B5;
        else
            page_size = HPDF_PAGE_SIZE_A4;

        // ピースのサイズ(mm)。
        double piece_size = wcstod(SETTING(IDC_PIECE_SIZE).c_str(), NULL) * 10;
        assert(piece_size > 0);

        // フォント名。
        string_t font_name = TEXT("MS-PGothic");

        // 外枠。
        auto frame_width_mm = _wtoi(SETTING(IDC_FRAME_WIDTH).c_str());
        auto frame_width_pixels = pixels_from_mm(frame_width_mm);

        // 線の太さ。
        double line_width = pixels_from_mm(0.8);

        // ピースの形状。
        bool abnormal = false;

        // カット線の色。
        double r_value, g_value, b_value;
        {
            auto text = SETTING(IDC_LINE_COLOR);
            auto ansi = ansi_from_wide(CP_ACP, text.c_str());
            uint32_t text_color = color_value_parse(ansi);
            text_color = color_value_fix(text_color);
            r_value = GetRValue(text_color) / 255.5;
            g_value = GetGValue(text_color) / 255.5;
            b_value = GetBValue(text_color) / 255.5;
        }

        // 背景画像。
        string_t background_image = SETTING(IDC_BACKGROUND_IMAGE);
        int background_width, background_height;
        HBITMAP hbmBackground = doLoadPic(background_image.c_str(), &background_width, &background_height);
        ::DeleteObject(hbmBackground);
        if (hbmBackground == NULL || !background_width || !background_height)
        {
            auto err_msg = ansi_from_wide(CP_ACP, doLoadString(IDS_CANT_LOAD_IMAGE));
            throw std::runtime_error(err_msg);
        }

        HPDF_Page page; // ページオブジェクト。
        HPDF_Font font; // フォントオブジェクト。
        double page_width, page_height; // ページサイズ。
        double content_x, content_y, content_width, content_height; // ページ内容の位置とサイズ。
        bool see_aspect = false;
        bool keep_aspect = false;
        for (INT iPage = 0; iPage < 3; ++iPage)
        {
            // ページを追加する。
            page = HPDF_AddPage(pdf);

            // ページサイズと用紙の向きを指定。
            HPDF_Page_SetSize(page, page_size, direction);

            // ページサイズ（ピクセル単位）。
            page_width = HPDF_Page_GetWidth(page);
            page_height = HPDF_Page_GetHeight(page);

            // アスペクト比。
            double aspect1 = page_width / page_height;
            double aspect2 = background_width / (double)background_height;
            if (!see_aspect)
            {
                see_aspect = true;
                if (aspect1 * 0.9 > aspect2 || aspect2 * 0.9 > aspect1)
                {
                    INT_PTR id = DialogBox(g_hInstance, MAKEINTRESOURCE(2), hwnd, AspectDlgProc);
                    switch (id)
                    {
                    case rad1:
                        break;
                    case rad2:
                        keep_aspect = true;
                        break;
                    case IDCANCEL:
                        return TEXT("");
                    default:
                        return TEXT("");
                    }
                }
            }

            // ページ内容の位置とサイズ。
            content_x = frame_width_pixels;
            content_y = frame_width_pixels;
            content_width = page_width - frame_width_pixels * 2;
            content_height = page_height - frame_width_pixels * 2;

            // 線の幅を指定。
            HPDF_Page_SetLineWidth(page, line_width);

            // 線の色を RGB で設定する。PDF では RGB 各値を [0,1] で指定することになっている。
            if (iPage == 0)
            {
                HPDF_Page_SetRGBStroke(page, r_value, g_value, b_value);
            }
            else
            {
                HPDF_Page_SetRGBStroke(page, 0, 0, 0);
            }

            /* 塗りつぶしの色を RGB で設定する。PDF では RGB 各値を [0,1] で指定することになっている。*/
            HPDF_Page_SetRGBFill(page, 0, 0, 0);

            // 行数と列数。
            int columns = (int)(mm_from_pixels(content_width) / piece_size);
            int rows = (int)(mm_from_pixels(content_height) / piece_size);
            if (rows <= 0)
                rows = 1;
            if (columns <= 0)
                columns = 1;
            double piece_width = content_width / columns;
            double piece_height = content_height / rows;
            if (0)
            {
                TCHAR szText[128];
                StringCchPrintf(szText, _countof(szText), TEXT("%f, %f"), piece_width, piece_height);
                MessageBox(hwnd, szText, NULL, 0);
            }

            // 乱数の種。
            int seed = _wtoi(SETTING(IDC_RANDOM_SEED).c_str());

            switch (iPage)
            {
            case 0:
                // 画像を描く。
                hpdf_draw_image(pdf, page, 0, 0, page_width, page_height, background_image, keep_aspect);
                // カット線を描画する。
                hpdf_draw_cut_lines(pdf, page, content_x, content_y, content_width, content_height, rows, columns, seed);
                break;
            case 1:
                // カット線を描画する。
                hpdf_draw_cut_lines(pdf, page, content_x, content_y, content_width, content_height, rows, columns, seed);
                break;
            case 2:
                // 画像を描く。
                hpdf_draw_image(pdf, page, 0, 0, page_width, page_height, background_image, keep_aspect);
                break;
            }

#ifndef NO_SHAREWARE
            // フォントを指定する。
            auto font_name_a = ansi_from_wide(CP932, font_name.c_str());
            font = HPDF_GetFont(pdf, font_name_a, "90ms-RKSJ-H");

            // シェアウェア未登録ならば、ロゴ文字列を描画する。
            if (!g_shareware.IsRegistered())
            {
                auto logo_a = ansi_from_wide(CP932, doLoadString(IDS_LOGO));
                double logo_x = content_x, logo_y = content_y;

                // フォントとフォントサイズを指定。
                HPDF_Page_SetFontAndSize(page, font, 16);

                // テキストを描画する。
                HPDF_Page_BeginText(page);
                {
                    HPDF_Page_TextOut(page, logo_x, logo_y, logo_a);
                }
                HPDF_Page_EndText(page);
            }
#endif
        }

        {
            // 現在の日時を取得する。
            SYSTEMTIME st;
            ::GetLocalTime(&st);

            // 出力ファイル名。
            TCHAR output_filename[MAX_PATH];
            StringCchPrintf(output_filename, _countof(output_filename),
                doLoadString(IDS_OUTPUT_NAME),
                st.wYear, st.wMonth, st.wDay, st.wHour, st.wMinute, st.wSecond);

            // PDFを一時ファイルに保存する。
            TempFile temp_file(TEXT("GN2"), TEXT(".pdf"));
            std::string temp_file_a = ansi_from_wide(CP_ACP, temp_file.make());
            HPDF_SaveToFile(pdf, temp_file_a.c_str());

            // デスクトップにファイルをコピー。
            TCHAR szPath[MAX_PATH];
            SHGetSpecialFolderPath(hwnd, szPath, CSIDL_DESKTOPDIRECTORY, FALSE);
            PathAppend(szPath, output_filename);
            if (!CopyFile(temp_file.get(), szPath, FALSE))
            {
                auto err_msg = ansi_from_wide(CP_ACP, doLoadString(IDS_COPYFILEFAILED));
                throw std::runtime_error(err_msg);
            }

            // 成功メッセージを表示。
            StringCchCopy(szPath, _countof(szPath), output_filename);
            TCHAR szText[MAX_PATH];
            StringCchPrintf(szText, _countof(szText), doLoadString(IDS_SUCCEEDED), szPath);
            ret = szText;
        }
    }
    catch (std::runtime_error& err)
    {
        // 失敗。
        auto wide = wide_from_ansi(CP_ACP, err.what());
        MessageBoxW(hwnd, wide, NULL, MB_ICONERROR);
        return TEXT("");
    }

    // PDFオブジェクトを解放する。
    HPDF_Free(pdf);

    return ret;
}

// WM_INITDIALOG
// ダイアログの初期化。
BOOL OnInitDialog(HWND hwnd, HWND hwndFocus, LPARAM lParam)
{
    // ユーザデータ。
    JigsawMake* pJM = (JigsawMake*)lParam;

    // ユーザーデータをウィンドウハンドルに関連付ける。
    SetWindowLongPtr(hwnd, GWLP_USERDATA, lParam);

    // gpimageライブラリを初期化する。
    gpimage_init();

    // ファイルドロップを受け付ける。
    DragAcceptFiles(hwnd, TRUE);

    // アプリの名前。
    LoadString(NULL, IDS_APPNAME, g_szAppName, _countof(g_szAppName));

    // アイコンの設定。
    g_hIcon = LoadIcon(g_hInstance, MAKEINTRESOURCE(1));
    g_hIconSm = (HICON)LoadImage(g_hInstance, MAKEINTRESOURCE(1), IMAGE_ICON,
        GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CYSMICON), 0);
    SendMessage(hwnd, WM_SETICON, ICON_BIG, (WPARAM)g_hIcon);
    SendMessage(hwnd, WM_SETICON, ICON_SMALL, (WPARAM)g_hIconSm);

    // 初期化。
    pJM->InitDialog(hwnd);

    // レジストリからデータを読み込む。
    pJM->DataFromReg(hwnd);

    // ダイアログにデータを設定。
    pJM->DialogFromData(hwnd);

    // Susie プラグインを読み込む。
    CHAR szPathA[MAX_PATH];
    GetModuleFileNameA(NULL, szPathA, _countof(szPathA));
    PathRemoveFileSpecA(szPathA);
    if (!g_susie.load(szPathA))
    {
        PathAppendA(szPathA, "plugins");
        g_susie.load(szPathA);
    }

    return TRUE;
}

// 「OK」ボタンが押された。
BOOL OnOK(HWND hwnd)
{
    // ユーザデータ。
    JigsawMake* pJM = (JigsawMake*)GetWindowLongPtr(hwnd, GWLP_USERDATA);

    // 「処理中...」とボタンに表示する。
    HWND hButton = GetDlgItem(hwnd, IDC_GENERATE);
    SetWindowText(hButton, doLoadString(IDS_PROCESSINGNOW));

    // ダイアログからデータを取得。
    if (!pJM->DataFromDialog(hwnd)) // 失敗。
    {
        // ボタンテキストを元に戻す。
        SetWindowText(hButton, doLoadString(IDS_GENERATE));

        return FALSE; // 失敗。
    }

    // 設定をレジストリに保存。
    pJM->RegFromData(hwnd);

    // メインディッシュ処理。
    string_t success = pJM->JustDoIt(hwnd);

    // ボタンテキストを元に戻す。
    SetWindowText(hButton, doLoadString(IDS_GENERATE));

    // 必要なら結果をメッセージボックスとして表示する。
    if (success.size())
    {
        MessageBox(hwnd, success.c_str(), g_szAppName, MB_ICONINFORMATION);
    }

    return TRUE; // 成功。
}

// 「設定の初期化」ボタン。
void OnEraseSettings(HWND hwnd)
{
    // ユーザーデータ。
    JigsawMake* pJM = (JigsawMake*)GetWindowLongPtr(hwnd, GWLP_USERDATA);

    // データをリセットする。
    pJM->Reset();

    // データからダイアログへ。
    pJM->DialogFromData(hwnd);

    // データからレジストリへ。
    pJM->RegFromData(hwnd);
}

// 「参照...」ボタン。
void OnBrowse(HWND hwnd)
{
    // テキストを取得する。前後の空白を取り除く。
    TCHAR szFile[MAX_PATH];
    GetDlgItemText(hwnd, IDC_BACKGROUND_IMAGE, szFile, _countof(szFile));
    str_trim(szFile);

    // ファイルでなければテキストをクリア。
    if (!PathFileExists(szFile))
    {
        szFile[0] = 0;
    }

    OPENFILENAME ofn = { sizeof(ofn), hwnd };

    // フィルター文字列をリソースから読み込む。
    string_t strFilter = doLoadString(IDS_OPENFILTER);

    // Susieプラグインのサポート。
    if (g_susie.is_loaded())
    {
        std::string additional = g_susie.get_filter();
        auto wide = wide_from_ansi(CP_ACP, additional.c_str());
        strFilter += doLoadString(IDS_SUSIE_IMAGES);
        strFilter += L" (";
        strFilter += wide;
        strFilter += L")|";
        strFilter += wide;
        strFilter += L"|";
    }

    // フィルター文字列をNULL区切りにする。
    for (auto& ch : strFilter)
    {
        if (ch == TEXT('|'))
            ch = 0;
    }

    // フィルター文字列を指定。
    ofn.lpstrFilter = strFilter.c_str();

    // ファイルを指定。
    ofn.lpstrFile = szFile;
    ofn.nMaxFile = _countof(szFile);

    // 初期フォルダを指定。
    TCHAR szDir[MAX_PATH];
    SHGetSpecialFolderPath(hwnd, szDir, CSIDL_MYPICTURES, FALSE);
    ofn.lpstrInitialDir = szDir;

    // タイトルを設定。
    string_t title = doLoadString(IDS_SET_BACKGROUND_IMAGE);
    ofn.lpstrTitle = title.c_str();

    ofn.Flags = OFN_EXPLORER | OFN_ENABLESIZING | OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST | OFN_HIDEREADONLY;
    ofn.lpstrDefExt = TEXT("JPG");

    // ユーザーにファイルを選択してもらう。
    if (::GetOpenFileName(&ofn))
    {
        // テキストを設定。
        ::SetDlgItemText(hwnd, IDC_BACKGROUND_IMAGE, szFile);
    }
}

// // 「カット線の色」ボタン。
void OnLineColorButton(HWND hwnd)
{
    // カット線の色を取得する。
    TCHAR szText[64];
    GetDlgItemText(hwnd, IDC_LINE_COLOR, szText, _countof(szText));
    StrTrim(szText, TEXT(" \t\r\n"));
    auto ansi = ansi_from_wide(CP_ACP, szText);
    auto color = color_value_parse(ansi);
    if (color == -1)
        color = 0;
    color = color_value_fix(color);

    static COLORREF custom_colors[16] = {
        RGB(255, 255, 255), RGB(255, 255, 255), RGB(255, 255, 255), RGB(255, 255, 255),
        RGB(255, 255, 255), RGB(255, 255, 255), RGB(255, 255, 255), RGB(255, 255, 255),
        RGB(255, 255, 255), RGB(255, 255, 255), RGB(255, 255, 255), RGB(255, 255, 255),
        RGB(255, 255, 255), RGB(255, 255, 255), RGB(255, 255, 255), RGB(255, 255, 255),
    };
    CHOOSECOLOR cc = { sizeof(cc), hwnd };
    cc.rgbResult = color;
    cc.lpCustColors = custom_colors;
    cc.Flags = CC_FULLOPEN | CC_RGBINIT;
    if (::ChooseColor(&cc))
    {
        StringCchPrintf(szText, _countof(szText), TEXT("#%02X%02X%02X"),
            GetRValue(cc.rgbResult),
            GetGValue(cc.rgbResult),
            GetBValue(cc.rgbResult)
        );
        SetDlgItemText(hwnd, IDC_LINE_COLOR, szText);
        InvalidateRect(GetDlgItem(hwnd, IDC_LINE_COLOR_BUTTON), NULL, TRUE);
    }
}

// WM_DRAWITEM
void OnDrawItem(HWND hwnd, const DRAWITEMSTRUCT * lpDrawItem)
{
    HWND hButton = ::GetDlgItem(hwnd, IDC_LINE_COLOR_BUTTON);
    if (lpDrawItem->hwndItem != hButton)
        return;

    // カット線の色を取得する。
    TCHAR szText[64];
    GetDlgItemText(hwnd, IDC_LINE_COLOR, szText, _countof(szText));
    StrTrim(szText, TEXT(" \t\r\n"));
    auto ansi = ansi_from_wide(CP_ACP, szText);
    auto text_color = color_value_parse(ansi);
    if (text_color == -1)
        text_color = 0;
    text_color = color_value_fix(text_color);

    BOOL bPressed = !!(lpDrawItem->itemState & ODS_CHECKED);

    // ボタンを描画する。
    RECT rcItem = lpDrawItem->rcItem;
    UINT uState = DFCS_BUTTONPUSH | DFCS_ADJUSTRECT;
    if (bPressed)
        uState |= DFCS_PUSHED;
    ::DrawFrameControl(lpDrawItem->hDC, &rcItem, DFC_BUTTON, uState);

    // 色ボタンの内側を描画する。
    HBRUSH hbr = ::CreateSolidBrush(text_color);
    ::FillRect(lpDrawItem->hDC, &rcItem, hbr);
    ::DeleteObject(hbr);
}

// WM_COMMAND
// コマンド。
void OnCommand(HWND hwnd, int id, HWND hwndCtl, UINT codeNotify)
{
    switch (id)
    {
    case IDC_GENERATE: // 「PDF生成」ボタン。
        OnOK(hwnd);
        break;
    case IDC_EXIT: // 「終了」ボタン。
        EndDialog(hwnd, id);
        break;
    case IDC_BROWSE: // 「参照...」ボタン。
        OnBrowse(hwnd);
        break;
    case IDC_ERASE_SETTINGS: // 「設定の初期化」ボタン。
        OnEraseSettings(hwnd);
        break;
    case IDC_LINE_COLOR: // 「カット線の色」テキストボックス。
        if (codeNotify == EN_CHANGE)
        {
            InvalidateRect(GetDlgItem(hwnd, IDC_LINE_COLOR_BUTTON), NULL, TRUE);
        }
        break;
    case IDC_LINE_COLOR_BUTTON: // 「カット線の色」ボタン。
        OnLineColorButton(hwnd);
        break;
    case stc1:
        // コンボボックスの前のラベルをクリックしたら、対応するコンボボックスにフォーカスを当てる。
        ::SetFocus(::GetDlgItem(hwnd, cmb1));
        break;
    case stc2:
        // コンボボックスの前のラベルをクリックしたら、対応するコンボボックスにフォーカスを当てる。
        ::SetFocus(::GetDlgItem(hwnd, cmb2));
        break;
    case stc3:
        // コンボボックスの前のラベルをクリックしたら、対応するコンボボックスにフォーカスを当てる。
        ::SetFocus(::GetDlgItem(hwnd, cmb3));
        break;
    case stc4:
        // コンボボックスの前のラベルをクリックしたら、対応するコンボボックスにフォーカスを当てる。
        ::SetFocus(::GetDlgItem(hwnd, cmb4));
        break;
    case stc5:
        // コンボボックスの前のラベルをクリックしたら、対応するコンボボックスにフォーカスを当てる。
        ::SetFocus(::GetDlgItem(hwnd, cmb5));
        break;
    case stc6:
        // テキストボックスの前のラベルをクリックしたら、対応するテキストボックスにフォーカスを当てる。
        {
            HWND hEdit = ::GetDlgItem(hwnd, edt1);
            Edit_SetSel(hEdit, 0, -1); // すべて選択。
            ::SetFocus(hEdit);
        }
        break;
    case stc7:
        // テキストボックスの前のラベルをクリックしたら、対応するテキストボックスにフォーカスを当てる。
        {
            HWND hEdit = ::GetDlgItem(hwnd, edt2);
            Edit_SetSel(hEdit, 0, -1); // すべて選択。
            ::SetFocus(hEdit);
        }
        break;
    case stc8:
        // テキストボックスの前のラベルをクリックしたら、対応するテキストボックスにフォーカスを当てる。
        {
            HWND hEdit = ::GetDlgItem(hwnd, edt8);
            Edit_SetSel(hEdit, 0, -1); // すべて選択。
            ::SetFocus(hEdit);
        }
        break;
    }
}

// WM_DROPFILES
// ファイルがドロップされた。
void OnDropFiles(HWND hwnd, HDROP hdrop)
{
    // ドロップ項目を取得する。
    TCHAR szFile[MAX_PATH];
    ::DragQueryFile(hdrop, 0, szFile, _countof(szFile));

    // 通常のファイルならば
    if (::PathFileExists(szFile) && !::PathIsDirectory(szFile))
    {
        // テキストボックスにテキストを設定。
        ::SetDlgItemText(hwnd, IDC_BACKGROUND_IMAGE, szFile);
    }

    // ドロップ完了。
    ::DragFinish(hdrop);
}

// WM_DESTROY
// ウィンドウが破棄された。
void OnDestroy(HWND hwnd)
{
    // アイコンを破棄。
    DestroyIcon(g_hIcon);
    DestroyIcon(g_hIconSm);
    g_hIcon = g_hIconSm = NULL;

    // gpimageライブラリを終了する。
    gpimage_exit();
}

// ダイアログプロシージャ。
INT_PTR CALLBACK
DialogProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    switch (uMsg)
    {
        HANDLE_MSG(hwnd, WM_INITDIALOG, OnInitDialog);
        HANDLE_MSG(hwnd, WM_DRAWITEM, OnDrawItem);
        HANDLE_MSG(hwnd, WM_COMMAND, OnCommand);
        HANDLE_MSG(hwnd, WM_DROPFILES, OnDropFiles);
        HANDLE_MSG(hwnd, WM_DESTROY, OnDestroy);
    }
    return 0;
}

// ガゾーナラベのメイン関数。
INT JigsawMake_Main(HINSTANCE hInstance, INT argc, LPTSTR *argv)
{
    // アプリのインスタンスを保持する。
    g_hInstance = hInstance;

    // 共通コントロール群を初期化する。
    InitCommonControls();

#ifndef NO_SHAREWARE
    // デバッガ―が有効、またはシェアウェアを開始できないときは
    if (IsDebuggerPresent() || !g_shareware.Start(NULL))
    {
        // 失敗。アプリケーションを終了する。
        return -1;
    }
#endif

    // ユーザーデータを保持する。
    JigsawMake gn(hInstance, argc, argv);

    // ユーザーデータをパラメータとしてダイアログを開く。
    DialogBoxParam(hInstance, MAKEINTRESOURCE(1), NULL, DialogProc, (LPARAM)&gn);

    // 正常終了。
    return 0;
}

// Windowsアプリのメイン関数。
INT WINAPI
WinMain(HINSTANCE   hInstance,
        HINSTANCE   hPrevInstance,
        LPSTR       lpCmdLine,
        INT         nCmdShow)
{
#ifdef UNICODE
    // wWinMainをサポートしていないコンパイラのために、コマンドラインの処理を行う。
    INT argc;
    LPWSTR *argv = CommandLineToArgvW(GetCommandLineW(), &argc);
    INT ret = JigsawMake_Main(hInstance, argc, argv);
    LocalFree(argv);
    return ret;
#else
    return JigsawMake_Main(hInstance, __argc, __argv);
#endif
}
