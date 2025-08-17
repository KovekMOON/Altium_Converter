#include <iostream>
#include <vector>
#include <string>
#include <cstring>
#include <fstream>
#include <filesystem>
#include <algorithm>
#include <regex>
#include <cctype>
#include <limits>
#include <unordered_map>
#include <unordered_set>
#include <optional>

#include <conio.h>
#include <windows.h>

namespace fs = std::filesystem;

// ---------- Settings ----------
static const std::string kFolder         = "For Conversion";     // входные файлы
static const std::string kComponentsDir  = "Components";         // база корпусов
static std::vector<std::string> kExts    = { /*".xls", ".xlsx", ".csv"*/ };
// -----------------------------

// Включаем ANSI-последовательности и UTF-8 один раз
static void enableVTMode() {
    HANDLE hOut = GetStdHandle(STD_OUTPUT_HANDLE);
    if (hOut != INVALID_HANDLE_VALUE) {
        DWORD mode = 0;
        if (GetConsoleMode(hOut, &mode)) {
            mode |= ENABLE_VIRTUAL_TERMINAL_PROCESSING;
            SetConsoleMode(hOut, mode);
        }
    }
    SetConsoleOutputCP(CP_UTF8);
    SetConsoleCP(CP_UTF8);
}

// утилиты
enum class Key { Up, Down, Enter, Other };

static inline void rstrip_cr(std::string& s) {
    if (!s.empty() && s.back() == '\r') s.pop_back();
}

static inline std::string ltrim(const std::string& s) {
    size_t i=0; while(i<s.size() && (s[i]==' '||s[i]=='\t')) ++i; return s.substr(i);
}
static inline std::string rtrim(const std::string& s) {
    if (s.empty()) return s; size_t j=s.size(); while(j>0 && (s[j-1]==' '||s[j-1]=='\t')) --j; return s.substr(0,j);
}
static inline std::string trim(const std::string& s) { return rtrim(ltrim(s)); }

static inline std::string tolower_copy(std::string s) {
    std::transform(s.begin(), s.end(), s.begin(),
        [](unsigned char c){ return std::tolower(c); });
    return s;
}

static std::string sanitize(std::string s) {
    if (s.empty()) return "UNKNOWN";
    for (char &c : s)
        if (!(std::isalnum((unsigned char)c) || c=='-' || c=='_' || c=='.')) c = '_';
    return s;
}

static bool hasExtension(const fs::path& p, const std::vector<std::string>& exts){
    if (exts.empty()) return true;
    std::string e = p.extension().string();
    std::transform(e.begin(), e.end(), e.begin(), ::tolower);
    for (auto x: exts) {
        std::string y = x; std::transform(y.begin(), y.end(), y.begin(), ::tolower);
        if (e == y) return true;
    }
    return false;
}

struct CompInfo {
    std::string standard;   // 2-й столбец
    bool to_delete = false; // 3-й столбец == "1"
};

// перечислить файлы базы
static std::vector<std::string> listComponentDbFiles() {
    std::vector<std::string> v;
    if (!fs::exists(kComponentsDir) || !fs::is_directory(kComponentsDir)) return v;
    for (auto& e : fs::directory_iterator(kComponentsDir)) {
        if (fs::is_regular_file(e) && e.path().extension() == ".csv")
            v.push_back(e.path().string());
    }
    std::sort(v.begin(), v.end());
    return v;
}

static std::vector<std::string> splitBySemicolon(const std::string& line) {
    std::vector<std::string> result; result.reserve(8);
    std::string cell;
    for (char c : line) {
        if (c == ';') { result.push_back(cell); cell.clear(); }
        else          { cell.push_back(c); }
    }
    result.push_back(cell);
    return result;
}

// разбор строки базы "nonstd;std;del"
static bool parseDbRow(const std::string& line, std::string& nonstd, CompInfo& info) {
    auto cells = splitBySemicolon(line);
    if (cells.size() < 3) return false;
    nonstd        = trim(cells[0]);
    info.standard = trim(cells[1]);
    std::string del = trim(cells[2]);
    info.to_delete = (del == "1");
    return !nonstd.empty();
}

// собрать индекс: key = lower(nonstd)
static std::unordered_map<std::string, CompInfo> buildComponentsIndexMap() {
    std::unordered_map<std::string, CompInfo> map;
    auto files = listComponentDbFiles();
    for (const auto& path : files) {
        std::ifstream in(path, std::ios::binary);
        if (!in) continue;
        std::string line; bool first = true;
        while (std::getline(in, line)) {
            rstrip_cr(line);
            if (first) { first = false; continue; } // пропускаем заголовок
            std::string nonstd; CompInfo info;
            if (parseDbRow(line, nonstd, info)) {
                map[tolower_copy(nonstd)] = info; // последнее определение побеждает
            }
        }
    }
    return map;
}

// аккуратное чтение 1-го столбца (до первого ';')
static std::string firstCellOfLine(const std::string& line) {
    size_t p = line.find(';');
    return trim( (p==std::string::npos)? line : line.substr(0,p) );
}

// построить индекс известных «нестандартных имён» из всех файлов базы
static std::unordered_set<std::string> buildComponentsIndex() {
    std::unordered_set<std::string> idx;
    auto files = listComponentDbFiles();
    for (const auto& path : files) {
        std::ifstream in(path, std::ios::binary);
        if (!in) continue;
        std::string line; bool first = true;
        while (std::getline(in, line)) {
            rstrip_cr(line);
            if (first) { first = false; continue; } // пропустить заголовок
            auto nonstd = firstCellOfLine(line);
            if (!nonstd.empty()) idx.insert(tolower_copy(nonstd));
        }
    }
    return idx;
}

// быстрое да/нет
static bool askYesNo(const std::string& question, bool defNo = true) {
    while (true) {
        std::cout << question << (defNo ? " [y/N]: " : " [Y/n]: ");
        std::string s; std::getline(std::cin, s);
        if (s.empty()) return !defNo;
        char c = (char)std::tolower((unsigned char)s[0]);
        if (c=='y') return true;
        if (c=='n') return false;
    }
}


// запись строки в CSV с ';' и кавычками по необходимости
static void writeCsvSemicolonRow(std::ofstream& f, const std::vector<std::string>& cells) {
    for (size_t i = 0; i < cells.size(); ++i) {
        const std::string& s = cells[i];
        bool need_quotes = s.find(';')!=std::string::npos || s.find('"')!=std::string::npos ||
                           s.find('\n')!=std::string::npos || s.find('\r')!=std::string::npos;
        if (need_quotes) {
            f << '"';
            for (char ch : s) f << (ch=='"' ? "\"\"" : std::string(1,ch));
            f << '"';
        } else {
            f << s;
        }
        if (i + 1 < cells.size()) f << ';';
    }
    f << "\n";
}



static std::vector<std::string> listFiles(const std::string& dir){
    std::vector<std::string> v;
    if(!fs::exists(dir) || !fs::is_directory(dir)) return v;
    for (const auto& ent: fs::directory_iterator(dir)) {
        if (fs::is_regular_file(ent.path()) && hasExtension(ent.path(), kExts))
            v.push_back(ent.path().string());
    }
    std::sort(v.begin(), v.end());
    return v;
}

static void renderMenu(const std::vector<std::string>& items, int selected, const std::string& title) {
    std::cout << "\x1b[H"; // Home
    std::cout << title << "\n\n";
    for (int i = 0; i < (int)items.size(); ++i) {
        std::cout << "\x1b[K";
        if (i == selected) std::cout << "\x1b[7m" << (i+1) << ". " << items[i] << "\x1b[0m\n";
        else               std::cout << (i+1) << ". " << items[i] << "\n";
    }
    std::cout << "\n\x1b[KUse Up/Down, Enter. (F5 = refresh)\n";
    std::cout.flush();
}

static int selectIndex(const std::vector<std::string>& items, const std::string& title) {
    std::cout << "\x1b[2J\x1b[H"; // clear once
    int sel = 0;
    renderMenu(items, sel, title);
    while (true) {
        int ch = _getch();
        if (ch == 13) return sel; // Enter
        if (ch == 0 || ch == 224) {
            int ch2 = _getch();
            if (ch2 == 72) { sel = (sel - 1 + (int)items.size()) % (int)items.size(); renderMenu(items, sel, title); } // ↑
            else if (ch2 == 80) { sel = (sel + 1) % (int)items.size(); renderMenu(items, sel, title); }                // ↓
            else if (ch2 == 63)  return -2; // F5 → обновить список
        }
    }
}

static void swallow_utf8_bom(std::ifstream& in) { // пропустить BOM UTF-8
    if (in.peek() == '\xEF') {
        char bom[3]; in.read(bom, 3);
        if (!(bom[0]=='\xEF' && bom[1]=='\xBB' && bom[2]=='\xBF')) {
            in.clear(); in.seekg(0);
        }
    }
}

// показать меню выбора файла базы (из Components)
static std::string pickComponentsFile(const std::string& forComponent) {
    auto files = listComponentDbFiles();
    if (files.empty()) return "";
    std::vector<std::string> items;
    for (auto& f : files) items.push_back(fs::path(f).filename().string());
    int idx = selectIndex(items, "=== Выберите CSV для: \"" + forComponent + "\" ===");
    if (idx < 0 || idx >= (int)files.size()) return "";
    return files[idx];
}

// предложить добавить компонент в выбранную базу
static std::optional<CompInfo>
maybeAddComponentToDb(const std::string& nonStandardName,
                      std::unordered_map<std::string, CompInfo>& idx /* онлайн-обновление */) {
    std::cout << "Компонент не найден в базе: \"" << nonStandardName << "\"\n";
    if (!askYesNo("Добавить этот компонент в базу?", true))
        return std::nullopt;

    // выбрать CSV в Components
    std::string dbPath = pickComponentsFile(nonStandardName);
    if (dbPath.empty()) {
        std::cout << "Файл базы не выбран. Пропуск.\n";
        return std::nullopt;
    }

    // ввести стандартное имя (по умолчанию = nonStandardName)
    std::cout << "Введите стандартное имя (Enter = оставить как есть):\n";
    std::cout << "  non-std: " << nonStandardName << "\n";
    std::cout << "  std    : ";
    std::string stdName;
    std::getline(std::cin, stdName);
    if (stdName.empty()) stdName = nonStandardName;

    // спросить флаг удаления
    bool del = askYesNo("Удалять этот элемент в будущем? (1 = да, 0 = нет)", false);
    std::string delFlag = del ? "1" : "0";

    // дописать строку в CSV
    fs::create_directories(fs::path(dbPath).parent_path());
    std::ofstream out(dbPath, std::ios::app | std::ios::binary);
    if (!out) {
        std::cerr << "Не удалось открыть для записи: " << dbPath << "\n";
        return std::nullopt;
    }
    writeCsvSemicolonRow(out, { nonStandardName, stdName, delFlag });

    // обновить индекс в памяти
    CompInfo info{ stdName, del };
    idx[tolower_copy(nonStandardName)] = info;

    std::cout << "Добавлено в: " << fs::path(dbPath).filename().string() << "\n";
    return info;
}


// -------- Нормализация домена --------
static std::string normalizeChipCapacitor(const std::string& in) {
    using std::regex; using std::regex_replace; using std::regex_constants::icase;
    std::string s = in;

    // 0402|0603|0805|1206|1210 - X5R|X7R - <VOLTAGE...> → корпус-<VOLTAGE...>
    static const regex re_x5x7(R"(^\s*(0402|0603|0805|1206|1210)\s*-\s*(X5R|X7R)\s*-\s*([0-9]+V-.*)$)", icase);
    if (std::regex_match(s, re_x5x7)) return regex_replace(s, re_x5x7, "$1-$3");

    // 0402|... - NP0 - <VOLTAGE...> → корпус-N<VOLTAGE...>
    static const regex re_np0(R"(^\s*(0402|0603|0805|1206|1210)\s*-\s*NP0\s*-\s*([0-9]+V-.*)$)", icase);
    if (std::regex_match(s, re_np0)) return regex_replace(s, re_np0, "$1-N$2");

    return s;
}

// Резисторы: <Pkg>-<Power W>-<Value>-<ppm>[-<tol%>] → убрать -ppm-, tol оставить
static std::string normalizeResistor(const std::string& in) {
    using std::regex; using std::smatch; using std::string;

    static const regex re(
        R"(^\s*([^- \t]+)\s*-\s*([0-9]+(?:\.[0-9]+)?|[0-9]+/[0-9]+)\s*W\s*-\s*(0R|[0-9]+(?:\.[0-9]+)?[RKM])\s*-\s*[0-9]+\s*ppm(?:\s*-\s*([0-9]+%))?\s*$)",
        std::regex::icase
    );
    smatch m;
    if (!std::regex_search(in, m, re)) return in;

    string out = m[1].str() + "-" + m[2].str() + "W-" + m[3].str();
    if (m[4].matched && !m[4].str().empty()) out += "-" + m[4].str();
    return out;
}

// Заменяет визуально похожие кириллические буквы на латинские (UTF-8 → ASCII)
static std::string fixCyrillicLetters(std::string s) {
    // Пары: UTF-8 кириллица → ASCII латиница
    static const std::pair<const char*, const char*> map[] = {
        // Верхний регистр
        {u8"А", "A"}, {u8"В", "B"}, {u8"Е", "E"}, {u8"К", "K"},
        {u8"М", "M"}, {u8"Н", "H"}, {u8"О", "O"}, {u8"Р", "P"},
        {u8"С", "C"}, {u8"Т", "T"}, {u8"У", "Y"}, {u8"Х", "X"},
        // Нижний регистр
        {u8"а", "a"}, {u8"в", "b"}, {u8"е", "e"}, {u8"к", "k"},
        {u8"м", "m"}, {u8"н", "n"}, {u8"о", "o"}, {u8"р", "p"},
        {u8"с", "c"}, {u8"т", "t"}, {u8"у", "y"}, {u8"х", "x"}
    };

    for (const auto& p : map) {
        const char* from = p.first;
        const char* to   = p.second;
        const size_t from_len = std::strlen(from);
        const size_t to_len   = std::strlen(to);

        size_t pos = 0;
        while ((pos = s.find(from, pos)) != std::string::npos) {
            s.replace(pos, from_len, to);
            pos += to_len;
        }
    }
    return s;
}

// Глобальная нормализация ячейки (ваши правила)
static std::string normalizeCell(const std::string& in){
    std::string s = in;

    // удалить (...) фрагменты
    static const std::regex paren_re(R"(\([^)]*\))");
    s = std::regex_replace(s, paren_re, "");

    // ',' → '.'
    for (char& c : s) if (c == ',') c = '.';

    // '?' → '-'
    for (char& c : s) if (c == '?') c = '-';

    // убрать пробелы перед тире
    static const std::regex dash_space_re(R"(\s+-)");
    s = std::regex_replace(s, dash_space_re, "-");

    // обрезать всё после '%' (сам знак оставить)
    if (size_t p = s.find('%'); p != std::string::npos) {
        s = s.substr(0, p+1);
    }

    // убрать всё до первой буквы/цифры/%
    size_t pos = 0;
    while (pos < s.size()) {
        unsigned char ch = static_cast<unsigned char>(s[pos]);
        if (std::isalnum(ch) || ch=='%') break;
        ++pos;
    }
    if (pos > 0) s.erase(0, pos);

    // если в исходнике было LESR — гарантировать его присутствие
    if (tolower_copy(in).find("lesr") != std::string::npos) {
        if (tolower_copy(s).find("lesr") == std::string::npos) {
            if (!s.empty() && s.back()=='%') s += "LESR";
            else                             s += " LESR";
        }
    }

    s = trim(s);

    // спец-правила домена
    s = normalizeChipCapacitor(s);
    s = normalizeResistor(s);
    s = fixCyrillicLetters(s);

    return s;
}

// ---------- Инициализация каталогов и базы корпусов ----------
int Initialisation() {
    enableVTMode();

    // создать каталоги
    for (const auto& folder : { "Settings", "For Conversion", "Converted", kComponentsDir.c_str(), "Documents" }) {
        fs::path dir = fs::current_path() / folder;
        if (!fs::exists(dir)) {
            if (!fs::create_directory(dir)) {
                std::cerr << "Failed to create: " << dir << "\n";
            }
        }
    }

    // файл со списком корпусов
    const std::string filename = "Settings/Packages.csv";
    if (!fs::exists(filename)) {
        std::ofstream file(filename, std::ios::binary);
    }

    std::ifstream infile(filename);
    if (!infile) {
        std::cerr << "Failed to open " << filename << " for read\n";
        return 1;
    }

    const std::string header = "Component_Name_Non_Standart;Component_Name_Standart;Delete_0_or_1\n";
    std::string line;
    while (std::getline(infile, line)) {
        if (line.empty()) continue;
        std::string fname = kComponentsDir + "/" + sanitize(line) + ".csv";
        if (fs::exists(fname)) continue;
        fs::create_directories(fs::path(fname).parent_path());
        std::ofstream out(fname, std::ios::binary);
        if (!out) { std::cerr << "Can't create " << fname << "\n"; continue; }
        out << header;
    }
    return 0;
}

static int processFile(const std::string& pathIn, const std::string& pathOut, bool verbose) {
    std::ifstream in(pathIn, std::ios::binary);
    if (!in) { std::cerr << "Не удалось открыть " << pathIn << "\n"; return 1; }
    swallow_utf8_bom(in);

    fs::create_directories(fs::path(pathOut).parent_path());
    std::ofstream out(pathOut, std::ios::binary);
    if (!out) { std::cerr << "Не удалось создать " << pathOut << "\n"; return 1; }

    // индекс базы — один на файл (чтобы не перечитывать для каждой строки)
    auto dbIndex = buildComponentsIndex();
    auto dbMap = buildComponentsIndexMap();

    std::string line;
    while (std::getline(in, line)) {
        rstrip_cr(line);
        auto cells = splitBySemicolon(line);

        // === твоя нормализация "каждый 5-й столбец" ===
        for (size_t i = 0; i < cells.size(); ++i) {
            if (((i + 1) % 5) == 0) {
                cells[i] = normalizeCell(cells[i]);
            }
        }

        // === правило поворота для C*/R* в 4-м столбце ===
        if (!cells.empty() && cells.size() > 3 && !cells[0].empty()) {
            char first = std::toupper(static_cast<unsigned char>(cells[0][0]));
            if (first == 'C' || first == 'R') {
                std::string rot = trim(cells[3]);
                if (rot == "180") cells[3] = "0";
                else if (rot == "270") cells[3] = "90";
            }
        }
        
        // вывод (опционально)
        if (verbose) {
            for (size_t i = 0; i < cells.size(); ++i)
                std::cout << "  [" << i << "] " << cells[i] << "\n";
            std::cout << "-------------------\n";
        }

        // === Сверка 5-го столбца с базой + интерактив для новых ===
        bool drop_line = false;
        if (cells.size() > 4) {
            std::string elem = trim(cells[4]);        // уже нормализован и латинизирован (normalizeCell в конце вызывает fixCyrillicLetters)
            std::string key  = tolower_copy(elem);

            auto it = dbMap.find(key);
            if (it != dbMap.end()) {
                const CompInfo& info = it->second;
                if (info.to_delete) {
                    drop_line = true;                 // удалить строку
                } else if (!info.standard.empty()) {
                    cells[4] = info.standard;         // заменить на стандарт
                }
            } else {
                // новый компонент → спросить, добавить ли
                if (auto added = maybeAddComponentToDb(elem, dbMap)) {
                    if (added->to_delete) {
                        drop_line = true;             // если пользователь отметил удалять — удаляем эту же строку
                    } else if (!added->standard.empty()) {
                        cells[4] = added->standard;   // иначе заменяем на новый стандарт
                    }
                } else {
                    // пользователь отказался добавлять — оставляем как есть
                }
            }
        }

if (drop_line) {
    if (verbose) std::cout << "↑ строка удалена по правилу базы\n";
    std::cout << "-------------------\n";
    continue; // НЕ записывать строку
}

        // запись строки в выходной CSV с ';'
        for (size_t i = 0; i < cells.size(); ++i) {
            cells[i] = fixCyrillicLetters(normalizeCell(cells[i]));
            out << cells[i];
            if (i + 1 < cells.size()) out << ';';
        }
        out << "\n";
    }
    return 0;
}



// -------- Конвертация --------
static int convertOne(const std::string& path) {
    std::cout << "\n Конвертация: " << path << "\n";
    const std::string outName = (fs::path("Converted") / fs::path(path).filename()).string();
    int rc = processFile(path, outName, /*verbose=*/true);
    if (rc == 0) std::cout << "Готово → " << outName << "\n";
    std::cout << "Нажмите Enter для возврата в меню...";
    std::cin.ignore(std::numeric_limits<std::streamsize>::max(), '\n');
    return rc;
}

static void convertAll(const std::vector<std::string>& files){
    if (files.empty()) {
        std::cout << "\n Нет файлов.\n Нажмите Enter...";
        std::cin.ignore(std::numeric_limits<std::streamsize>::max(), '\n');
        return;
    }
    std::cout << "\n Всего файлов: " << files.size() << "\n";
    size_t ok = 0;
    for (const auto& f : files) {
        const std::string outName = (fs::path("Converted") / fs::path(f).filename()).string();
        std::cout << "- " << fs::path(f).filename().string() << " → " << outName << "\n";
        int rc = processFile(f, outName, /*verbose=*/false);
        if (rc == 0) ++ok;
    }
    std::cout << "Готово: " << ok << "/" << files.size() << " успешно.\n"
              << "Нажмите Enter для возврата в меню...";
    std::cin.ignore(std::numeric_limits<std::streamsize>::max(), '\n');
}

// -------- Меню конвертации --------
static int selectAndConvert() {
    while (true) {
        auto files = listFiles(kFolder);

        std::vector<std::string> items;
        items.reserve(files.size() + 2);
        for (const auto& p: files) items.push_back(fs::path(p).filename().string());
        items.push_back("[Convert ALL]");
        items.push_back("[Exit]");

        int idx = selectIndex(items, "=== Files in \"" + kFolder + "\" ===");
        if (idx == -2) continue; // F5

        if (idx == (int)items.size() - 1) { std::cout << "Exit.\n"; return 0; }
        if (idx == (int)items.size() - 2) { convertAll(files); continue; }

        if (idx >= 0 && idx < (int)files.size()) {
            convertOne(files[idx]);
        }
    }
}

// -------- Заглушки других пунктов --------
static void compareWithDocs() {
    std::cout << "[Compare with documentation] недоступно в этой версии.\n";
    std::cout << "Нажмите Enter..."; std::cin.ignore(std::numeric_limits<std::streamsize>::max(), '\n');
}
static void databaseOfComponents() {
    std::cout << "[Database of components] недоступно в этой версии.\n";
    std::cout << "Нажмите Enter..."; std::cin.ignore(std::numeric_limits<std::streamsize>::max(), '\n');
}

// -------- main --------
int main() {
    if (Initialisation() != 0) return 1;  // важна успешная инициализация

    enableVTMode(); // ANSI + UTF-8

    std::vector<std::string> items = {
        "Convert altium",
        "Compare with documentation",
        "Database of components",
        "Exit"
    };

    int choice = selectIndex(items, "=== Main Menu ===");
    switch (choice) {
        case 0: selectAndConvert();    break;
        case 1: compareWithDocs();     break;
        case 2: databaseOfComponents();break;
        case 3: std::cout << "Exit.\n"; break;
    }
    return 0;
}

