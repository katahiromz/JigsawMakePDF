// デジタルジグソーメーカー by katahiromz
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
#include "resource.h"       // リソースIDの定義ヘッダ。

// シェアウェア情報。
#ifndef NO_SHAREWARE
    #include "Shareware.hpp"

    SW_Shareware g_shareware(
        /* company registry key */      TEXT("Katayama Hirofumi MZ"),
        /* application registry key */  TEXT("DigitalJigsawMaker"),
        /* password hash */
        "e218f83f070a186f886c6dc82bd7ecf3d6c3ea4224fd7d213aa06e9c9713b395",
        /* trial days */                10,
        /* salt string */               "mJpDxx2D",
        /* version string */            "0.0.0");
#endif

#define UTF8_SUPPORT // UTF-8サポート。

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
class JigsawMaker
{
public:
    HINSTANCE m_hInstance;
    INT m_argc;
    LPTSTR *m_argv;
    HBITMAP m_hbmPreview;
    std::map<string_t, string_t> m_settings;
    std::vector<string_t> m_list;
    std::vector<FONT_ENTRY> m_font_map;

    // コンストラクタ。
    JigsawMaker(HINSTANCE hInstance, INT argc, LPTSTR *argv);

    // デストラクタ。
    ~JigsawMaker()
    {
        ::DeleteObject(m_hbmPreview);
    }

    // フォントマップを読み込む。
    BOOL LoadFontMap();
    // データをリセットする。
    void Reset();
    // ダイアログを初期化する。
    void InitDialog(HWND hwnd);
    // ダイアログからデータへ。
    BOOL DataFromDialog(HWND hwnd, BOOL bList);
    // データからダイアログへ。
    BOOL DialogFromData(HWND hwnd, BOOL bList);
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
doLoadPic(LPCWSTR filename, int* width = NULL, int* height = NULL,
          FILETIME* pftCreated = NULL, FILETIME* pftModified = NULL)
{
    // GDI+などで読み込みを試みる。
    HBITMAP hbm = gpimage_load(filename, width, height, pftCreated, pftModified);
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
        gpimage_load_datetime(filename, pftCreated, pftModified);
        return hbm;
    }

    return NULL; // 失敗。
}

// フォントマップを読み込む。
BOOL JigsawMaker::LoadFontMap()
{
    // 初期化する。
    m_font_map.clear();

    // ローカルファイルの「fontmap.dat」を探す。
    auto filename = findLocalFile(TEXT("fontmap.dat"));
    if (filename == NULL)
        return FALSE; // 見つからなかった。

    // ファイル「fontmap.dat」を開く。
    if (FILE *fp = _tfopen(filename, TEXT("rb")))
    {
        // 一行ずつ読み込む。
        char buf[512];
        while (fgets(buf, _countof(buf), fp))
        {
            // UTF-8文字列をワイド文字列に変換する。
            WCHAR szText[512];
            MultiByteToWideChar(CP_UTF8, 0, buf, -1, szText, _countof(szText));

            // 前後の空白を取り除く。
            str_trim(szText);

            // 行コメントを削除する。
            if (auto pch = wcschr(szText, L';'))
            {
                *pch = 0;
            }

            // もう一度前後の空白を取り除く。
            str_trim(szText);

            // 「=」を探す。
            if (auto pch = wcschr(szText, L'='))
            {
                // 文字列を切り分ける。
                *pch++ = 0;
                auto font_name = szText;
                auto font_file = pch;

                // 前後の空白を取り除く。
                str_trim(font_name);
                str_trim(font_file);

                // 「,」があればインデックスを読み込み、切り分ける。
                pch = wcschr(pch, L',');
                int index = -1;
                if (pch)
                {
                    *pch++ = 0;
                    index = _wtoi(pch);
                }

                // さらに前後の空白を取り除く。
                str_trim(font_name);
                str_trim(font_file);

                // フォントファイルのパスファイル名を構築する。
                TCHAR font_pathname[MAX_PATH];
                GetWindowsDirectory(font_pathname, _countof(font_pathname));
                PathAppend(font_pathname, TEXT("Fonts"));
                PathAppend(font_pathname, font_file);

                // パスファイル名が存在するか？
                if (PathFileExists(font_pathname))
                {
                    // 存在すれば、フォントのエントリーを追加。
                    FONT_ENTRY entry;
                    entry.m_font_name = font_name;
                    entry.m_pathname = font_pathname;
                    entry.m_index = index;
                    m_font_map.push_back(entry);
                }
            }
        }

        // ファイルを閉じる。
        fclose(fp);
    }

    return m_font_map.size() > 0;
}

// コンストラクタ。
JigsawMaker::JigsawMaker(HINSTANCE hInstance, INT argc, LPTSTR *argv)
    : m_hInstance(hInstance)
    , m_argc(argc)
    , m_argv(argv)
    , m_hbmPreview(NULL)
{
    // データをリセットする。
    Reset();

    // フォントマップを読み込む。
    LoadFontMap();

    // コマンドライン引数をリストに追加。
    for (INT i = 1; i < m_argc; ++i)
    {
        // 有効な画像ファイルかを確認して追加。
        if (isValidImageFile(m_argv[i]))
            m_list.push_back(m_argv[i]);
    }
}

// 既定値。
#define IDC_PAGE_SIZE_DEFAULT doLoadString(IDS_A4)
#define IDC_PAGE_DIRECTION_DEFAULT doLoadString(IDS_LANDSCAPE)

// データをリセットする。
void JigsawMaker::Reset()
{
#define SETTING(id) m_settings[TEXT(#id)]
    SETTING(IDC_PAGE_SIZE) = IDC_PAGE_SIZE_DEFAULT;
    SETTING(IDC_PAGE_DIRECTION) = IDC_PAGE_DIRECTION_DEFAULT;
}

// ダイアログを初期化する。
void JigsawMaker::InitDialog(HWND hwnd)
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
}

// ダイアログからデータへ。
BOOL JigsawMaker::DataFromDialog(HWND hwnd)
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
#undef GET_COMBO_DATA

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
BOOL JigsawMaker::DialogFromData(HWND hwnd)
{
    // コンボボックスへデータを設定する。
#define SET_COMBO_DATA(id) \
    setComboText(hwnd, (id), m_settings[TEXT(#id)].c_str());
    SET_COMBO_DATA(IDC_PAGE_SIZE);
    SET_COMBO_DATA(IDC_PAGE_DIRECTION);
#undef SET_COMBO_DATA

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
BOOL JigsawMaker::DataFromReg(HWND hwnd)
{
    // ソフト固有のレジストリキーを開く。
    HKEY hKey;
    RegOpenKeyEx(HKEY_CURRENT_USER, TEXT("Software\\Katayama Hirofumi MZ\\DigitalJigsawMaker"), 0, KEY_READ, &hKey);
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
#undef GET_REG_DATA

    // レジストリキーを閉じる。
    RegCloseKey(hKey);
    return TRUE; // 成功。
}

// データからレジストリへ。
BOOL JigsawMaker::RegFromData(HWND hwnd)
{
    HKEY hCompanyKey = NULL, hAppKey = NULL;

    // 会社固有のレジストリキーを作成または開く。
    RegCreateKey(HKEY_CURRENT_USER, TEXT("Software\\Katayama Hirofumi MZ"), &hCompanyKey);
    if (hCompanyKey == NULL)
        return FALSE; // 失敗。

    // ソフト固有のレジストリキーを作成または開く。
    RegCreateKey(hCompanyKey, TEXT("DigitalJigsawMaker"), &hAppKey);
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
bool hpdf_draw_image(HPDF_Doc pdf, HPDF_Page page, double x, double y, double width, double height, const string_t& filename,
                     DWORD max_data_size, bool dont_resize_small)
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
        if (file_size == 0)
            break; // 失敗。

        // ファイルサイズは妥当か？
        if (max_data_size == 0 || file_size <= max_data_size)
            break; // 妥当。

        // 画像サイズを縮小した画像を作成する。ゼロにならないように補正する。
        image_width /= 2;
        image_height /= 2;
        if (image_width <= 0)
            image_width = 1;
        if (image_height <= 0)
            image_height = 1;
        HBITMAP hbmNew = gpimage_resize(hbm, image_width, image_height);
        if (hbmNew == NULL)
        {
            file_size = 0;
            break;
        }

        // hbmを縮小した画像に置き換える。
        ::DeleteObject(hbm);
        hbm = hbmNew;

        // 一時ファイルを削除。
        tempfile.erase();
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

    // 画像サイズとアスペクト比に従って処理を行う。
    double aspect1 = (double)image_width / (double)image_height;
    double aspect2 = width / height;
    double stretch_width, stretch_height;
    if (image_width <= width && image_height <= height && dont_resize_small)
    {
        // 小さい画像を拡大しない。
        stretch_width = image_width;
        stretch_height = image_height;
    }
    else
    {
        // 大きい画像はアスペクト比に従ってセルいっぱいに縮小する。
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

// メインディッシュ処理。
string_t JigsawMaker::JustDoIt(HWND hwnd)
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

#ifdef UTF8_SUPPORT
        // エンコーディング "UTF-8" が利用可能に？？？
        HPDF_UseUTFEncodings(pdf);
        HPDF_SetCurrentEncoder(pdf, "UTF-8");
#endif

        // 日本語フォントの MS-(P)Mincyo, MS-(P)Gothic が利用可能となる
        //HPDF_UseJPFonts(pdf);

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

        // フォント名。
        string_t font_name;
        if (m_font_map.size())
        {
            for (auto& entry : m_font_map)
            {
                if (entry.m_font_name != SETTING(IDC_FONT_NAME))
                    continue;

                auto ansi = ansi_from_wide(CP_ACP, entry.m_pathname.c_str());
                if (entry.m_index != -1)
                {
                    std::string font_name_a = HPDF_LoadTTFontFromFile2(pdf, ansi, entry.m_index, HPDF_TRUE);
                    font_name = wide_from_ansi(CP_UTF8, font_name_a.c_str());
                }
                else
                {
                    std::string font_name_a = HPDF_LoadTTFontFromFile(pdf, ansi, HPDF_TRUE);
                    font_name = wide_from_ansi(CP_UTF8, font_name_a.c_str());
                }
            }
        }
        else
        {
            if (SETTING(IDC_FONT_NAME) == doLoadString(IDS_FONT_01))
                font_name = TEXT("MS-PGothic");
            else if (SETTING(IDC_FONT_NAME) == doLoadString(IDS_FONT_02))
                font_name = TEXT("MS-PMincho");
            else if (SETTING(IDC_FONT_NAME) == doLoadString(IDS_FONT_03))
                font_name = TEXT("MS-Gothic");
            else if (SETTING(IDC_FONT_NAME) == doLoadString(IDS_FONT_04))
                font_name = TEXT("MS-Mincho");
            else
                font_name = TEXT("MS-PGothic");
        }

        HPDF_Page page; // ページオブジェクト。
        HPDF_Font font; // フォントオブジェクト。
        INT iColumn = 0, iRow = 0, iPage = 0; // セル位置とページ番号。
        INT cItems = INT(m_list.size()); // 項目数。
        INT cPages = (cItems + (columns * rows) - 1) / (columns * rows); // ページ総数。
        double page_width, page_height; // ページサイズ。
        double content_x, content_y, content_width, content_height; // ページ内容の位置とサイズ。
        for (INT iItem = 0; iItem < cItems; ++iItem)
        {
            if (iColumn == 0 && iRow == 0) // 項目がページの最初ならば
            {
                // ページを追加する。
                page = HPDF_AddPage(pdf);

                // ページサイズと用紙の向きを指定。
                HPDF_Page_SetSize(page, page_size, direction);

                // ページサイズ（ピクセル単位）。
                page_width = HPDF_Page_GetWidth(page);
                page_height = HPDF_Page_GetHeight(page);

                // ページ内容の位置とサイズ。
                content_x = margin;
                content_y = margin;
                content_width = page_width - margin * 2;
                content_height = page_height - margin * 2;

                // 線の幅を指定。
                HPDF_Page_SetLineWidth(page, border_width);

                // 線の色を RGB で設定する。PDF では RGB 各値を [0,1] で指定することになっている。
                HPDF_Page_SetRGBStroke(page, 0, 0, 0);

                /* 塗りつぶしの色を RGB で設定する。PDF では RGB 各値を [0,1] で指定することになっている。*/
                HPDF_Page_SetRGBFill(page, 0, 0, 0);

                // フォントを指定する。
                auto font_name_a = ansi_from_wide(CP932, font_name.c_str());
#ifdef UTF8_SUPPORT
                font = HPDF_GetFont(pdf, font_name_a, "UTF-8");
#else
                font = HPDF_GetFont(pdf, font_name_a, "90ms-RKSJ-H");
#endif

#ifndef NO_SHAREWARE
                // シェアウェア未登録ならば、ロゴ文字列を描画する。
                if (!g_shareware.IsRegistered())
                {
#ifdef UTF8_SUPPORT
                    auto logo_a = ansi_from_wide(CP_UTF8, doLoadString(IDS_LOGO));
#else
                    auto logo_a = ansi_from_wide(CP932, doLoadString(IDS_LOGO));
#endif
                    double logo_x = content_x, logo_y = content_y;

                    // フォントとフォントサイズを指定。
                    HPDF_Page_SetFontAndSize(page, font, footer_height);

                    // テキストを描画する。
                    HPDF_Page_BeginText(page);
                    {
                        HPDF_Page_TextOut(page, logo_x, logo_y, logo_a);
                    }
                    HPDF_Page_EndText(page);
                }
#endif
                // ヘッダー（ページ見出し）を描画する。
                if (header.size())
                {
                    string_t text = header;
                    substitute_tags(hwnd, text, m_list[iItem], iItem, cItems, iPage, cPages);
#ifdef UTF8_SUPPORT
                    auto header_text_a = ansi_from_wide(CP_UTF8, text.c_str());
#else
                    auto header_text_a = ansi_from_wide(CP932, text.c_str());
#endif
                    hpdf_draw_text(page, font, footer_height, header_text_a,
                                   content_x, content_y + content_height - footer_height,
                                   content_width, footer_height);
                    // ヘッダーの分だけページ内容のサイズを縮小する。
                    content_height -= footer_height;
                }

                // フッター（ページ番号）を描画する。
                if (page_numbers || footer.size())
                {
                    string_t text = footer;
                    substitute_tags(hwnd, text, m_list[iItem], iItem, cItems, iPage, cPages);
#ifdef UTF8_SUPPORT
                    auto footer_text_a = ansi_from_wide(CP_UTF8, text.c_str());
#else
                    auto footer_text_a = ansi_from_wide(CP932, text.c_str());
#endif
                    hpdf_draw_text(page, font, footer_height, footer_text_a,
                                   content_x, content_y,
                                   content_width, footer_height);
                    // フッターの分だけページ内容のサイズを縮小する。
                    content_y += footer_height;
                    content_height -= footer_height;
                }

                // 枠線を描く。
                if (draw_border)
                {
                    hpdf_draw_box(page, content_x, content_y, content_width, content_height);
                }
            }

            // セルの位置とサイズ。
            double cell_width = content_width / columns;
            double cell_height = content_height / rows;
            double cell_x = content_x + cell_width * iColumn;
            double cell_y = content_y + cell_height * (rows - iRow - 1);

            // セルの枠線を描く。
            if (draw_border)
            {
                hpdf_draw_box(page, cell_x, cell_y, cell_width, cell_height);

                // 枠線の分だけセルを縮小。
                cell_x += border_width;
                cell_y += border_width;
                cell_width -= border_width * 2;
                cell_height -= border_width * 2;
            }

            // テキストを描画する。
            if (image_title.size())
            {
                // タグを規則に従って置き換える。
                string_t text = image_title;
                substitute_tags(hwnd, text, m_list[iItem], iItem, cItems, iPage, cPages);

                // ANSI文字列に変換してテキストを描画する。
#ifdef UTF8_SUPPORT
                auto text_a = ansi_from_wide(CP_UTF8, text.c_str());
#else
                auto text_a = ansi_from_wide(CP932, text.c_str());
#endif
                hpdf_draw_text(page, font, font_size, text_a, cell_x, cell_y, cell_width, font_size);

                // セルのサイズを縮小する。
                cell_y += font_size;
                cell_height -= font_size;
            }

            // 画像を描く。
            hpdf_draw_image(pdf, page, cell_x, cell_y, cell_width, cell_height, m_list[iItem],
                            max_data_size, dont_resize_small);

            // 次のセルに進む。必要ならばページ番号を進める。
            ++iColumn;
            if (iColumn == columns)
            {
                iColumn = 0;
                ++iRow;
                if (iRow == rows)
                {
                    iRow = 0;
                    ++iPage;
                }
            }
        }

        {
            // 出力ファイル名のタグを置き換える。
            string_t text = output_name;
            substitute_tags(hwnd, text, m_list[0], 0, cItems, 0, cPages, true);

            // ファイル名に使えない文字を置き換える。
            validate_filename(text);

            // PDFを一時ファイルに保存する。
            TempFile temp_file(TEXT("GN2"), TEXT(".pdf"));
            std::string temp_file_a = ansi_from_wide(CP_ACP, temp_file.make());
            HPDF_SaveToFile(pdf, temp_file_a.c_str());

            // デスクトップにファイルをコピー。
            TCHAR szPath[MAX_PATH];
            SHGetSpecialFolderPath(hwnd, szPath, CSIDL_DESKTOPDIRECTORY, FALSE);
            PathAppend(szPath, text.c_str());
            StringCchCat(szPath, _countof(szPath), TEXT(".pdf"));
            if (!CopyFile(temp_file.get(), szPath, FALSE))
            {
                auto err_msg = ansi_from_wide(CP_ACP, doLoadString(IDS_COPYFILEFAILED));
                throw std::runtime_error(err_msg);
            }

            // 成功メッセージを表示。
            StringCchCopy(szPath, _countof(szPath), text.c_str());
            StringCchCat(szPath, _countof(szPath), TEXT(".pdf"));
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
    JigsawMaker* pJM = (JigsawMaker*)lParam;

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
    pJM->DialogFromData(hwnd, FALSE);

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
    JigsawMaker* pJM = (JigsawMaker*)GetWindowLongPtr(hwnd, GWLP_USERDATA);

    // 「処理中...」とボタンに表示する。
    HWND hButton = GetDlgItem(hwnd, IDC_GENERATE);
    SetWindowText(hButton, doLoadString(IDS_PROCESSINGNOW));

    // ダイアログからデータを取得。
    if (!pJM->DataFromDialog(hwnd, TRUE)) // 失敗。
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
    JigsawMaker* pJM = (JigsawMaker*)GetWindowLongPtr(hwnd, GWLP_USERDATA);

    // データをリセットする。
    pJM->Reset();

    // データからダイアログへ。
    pJM->DialogFromData(hwnd, FALSE);

    // データからレジストリへ。
    pJM->RegFromData(hwnd);
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
    }
}

// WM_DROPFILES
// ファイルがドロップされた。
void OnDropFiles(HWND hwnd, HDROP hdrop)
{
    // ドロップ完了。
    DragFinish(hdrop);
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
        HANDLE_MSG(hwnd, WM_COMMAND, OnCommand);
        HANDLE_MSG(hwnd, WM_DROPFILES, OnDropFiles);
        HANDLE_MSG(hwnd, WM_DESTROY, OnDestroy);
    }
    return 0;
}

// ガゾーナラベのメイン関数。
INT JigsawMaker_Main(HINSTANCE hInstance, INT argc, LPTSTR *argv)
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
    JigsawMaker gn(hInstance, argc, argv);

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
    INT ret = JigsawMaker_Main(hInstance, argc, argv);
    LocalFree(argv);
    return ret;
#else
    return JigsawMaker_Main(hInstance, __argc, __argv);
#endif
}
