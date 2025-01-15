#include <iostream>
#include <vector>
#include <windows.h>
#include <sql.h>
#include <sqlext.h>
#include <xlsxwriter.h>
#include <fstream>
#define pause system("pause")
#define cls system("cls")

bool fileExist(std::string filename) {
    std::ifstream file(filename);
    return file.good(); // Trả về true nếu file tồn tại
}
void checkError(SQLRETURN retCode, SQLHANDLE handle, SQLSMALLINT type) {
    if (retCode != SQL_SUCCESS && retCode != SQL_SUCCESS_WITH_INFO) {
        SQLWCHAR sqlState[1024];
        SQLWCHAR message[1024];
        SQLGetDiagRec(type, handle, 1, sqlState, nullptr, message, 1024, nullptr);
        std::cerr << "SQL Error: " << message << "\n";
        exit(-1);
    }
}
void exportDataToExcel(SQLRETURN retCode, SQLHSTMT hStmt, SQLHDBC hDbc,const std::string& filename) {

    // Tạo workbook và worksheet
    lxw_workbook* workbook = workbook_new(filename.c_str());
    lxw_worksheet* worksheet = workbook_add_worksheet(workbook, NULL);

    // Tạo định dạng cho tiêu đề (In đậm, nền màu vàng, căn giữa)
    lxw_format* header_format = workbook_add_format(workbook);
    format_set_bold(header_format);
    format_set_bg_color(header_format, LXW_COLOR_YELLOW);
    format_set_align(header_format, LXW_ALIGN_CENTER);

    // Tạo định dạng cho dữ liệu (Chữ xanh, căn trái)
    lxw_format* data_format = workbook_add_format(workbook);
    format_set_font_color(data_format, LXW_COLOR_BLUE);
    format_set_align(data_format, LXW_ALIGN_LEFT);

    // Viết dữ liệu với định dạng
    worksheet_write_string(worksheet, 0, 0, "ID", header_format);
    worksheet_write_string(worksheet, 0, 1, "NAME", header_format);
    worksheet_write_string(worksheet, 0, 2, "AGE", header_format);

    worksheet_write_string(worksheet, 1, 0, "Data 1", data_format);
    worksheet_write_string(worksheet, 1, 1, "Data 2", data_format);

    // Tạo định dạng số (2 chữ số thập phân)
    lxw_format* number_format = workbook_add_format(workbook);
    format_set_num_format(number_format, "0");

    // Ghi dữ liệu số với định dạng
    worksheet_write_number(worksheet, 2, 0, 1234.567, number_format);
    worksheet_write_number(worksheet, 2, 1, 9876.543, number_format);

    // Prepare the query
    SQLWCHAR query[] = L"SELECT * FROM SinhVien ";
    retCode = SQLPrepare(hStmt, query, SQL_NTS);

    checkError(retCode, hStmt, SQL_HANDLE_STMT);
    // Execute the query
    retCode = SQLExecute(hStmt);
    checkError(retCode, hStmt, SQL_HANDLE_STMT);

    // Fetch and print results
    SQLWCHAR name[51];  // Changed from SQLCHAR to SQLWCHAR
    SQLINTEGER id;
    SQLINTEGER age;
    int row = 1;
    while (SQLFetch(hStmt) == SQL_SUCCESS) {
        SQLGetData(hStmt, 1, SQL_C_SLONG, &id, 0, nullptr);
        SQLGetData(hStmt, 2, SQL_C_WCHAR, name, sizeof(name), nullptr);  // Changed to SQL_C_WCHAR
        SQLGetData(hStmt, 3, SQL_C_SLONG, &age, 0, nullptr);
        worksheet_write_number(worksheet, row, 0, id, number_format);
        worksheet_write_number(worksheet, row, 2, age, number_format);
        int len = wcslen(name) + 1;
        char* convertName = new char[len];

        size_t converted = 0;  // Số lượng ký tự đã chuyển đổi
        wcstombs_s(&converted, convertName, len, name, _TRUNCATE);
        worksheet_write_string(worksheet, row, 1, convertName, data_format);
        row++;
    }
    // Free the statement handle after data fetching
    retCode = SQLFreeStmt(hStmt, SQL_CLOSE); // Or SQL_UNBIND if you're not closing
    checkError(retCode, hStmt, SQL_HANDLE_STMT);

    // Đóng workbook
    workbook_close(workbook);
    std::cout << "Export Success\n";
}

void GetData(SQLRETURN retCode, SQLHSTMT hStmt, SQLHDBC hDbc) {
    // Prepare the query
    SQLWCHAR query[] = L"SELECT * FROM SinhVien ";
    retCode = SQLPrepare(hStmt, query, SQL_NTS);

    checkError(retCode, hStmt, SQL_HANDLE_STMT);
    // Execute the query
    retCode = SQLExecute(hStmt);
    checkError(retCode, hStmt, SQL_HANDLE_STMT);

    // Fetch and print results
    SQLWCHAR name[51];  // Changed from SQLCHAR to SQLWCHAR
    SQLINTEGER id;
    SQLINTEGER age;
    while (SQLFetch(hStmt) == SQL_SUCCESS) {
        SQLGetData(hStmt, 1, SQL_C_SLONG, &id, 0, nullptr);
        SQLGetData(hStmt, 2, SQL_C_WCHAR, name, sizeof(name), nullptr);  // Changed to SQL_C_WCHAR
        SQLGetData(hStmt, 3, SQL_C_SLONG, &age, 0, nullptr);

        std::wcout << L"ID: " << id << L" NAME: " << name << L" AGE: " << age << std::endl;
    }
    // Free the statement handle after data fetching
    retCode = SQLFreeStmt(hStmt, SQL_CLOSE); // Or SQL_UNBIND if you're not closing
    checkError(retCode, hStmt, SQL_HANDLE_STMT);
}

void Insert(SQLRETURN retCode, SQLHSTMT hStmt, SQLHDBC hDbc, SQLINTEGER id, SQLWCHAR name[], SQLINTEGER age) {
    SQLWCHAR query[] = L"INSERT INTO SinhVien values(?,?,?);";
    retCode = SQLPrepare(hStmt, query, SQL_NTS);

    if (retCode != SQL_SUCCESS && retCode != SQL_SUCCESS_WITH_INFO) {
        std::wcout << L"Error in SQLPrepare.\n";
        return;
    }

    // Bind parameters
    SQLBindParameter(hStmt, 1, SQL_PARAM_INPUT, SQL_C_SLONG, SQL_INTEGER, 0, 0, &id, 0, nullptr);
    SQLBindParameter(hStmt, 2, SQL_PARAM_INPUT, SQL_C_WCHAR, SQL_WCHAR, 50, 0, name, sizeof(name), nullptr);  // Corrected the binding size
    SQLBindParameter(hStmt, 3, SQL_PARAM_INPUT, SQL_C_SLONG, SQL_INTEGER, 0, 0, &age, 0, nullptr);

    // Execute the prepared statement
    retCode = SQLExecute(hStmt);
    if (retCode == SQL_SUCCESS || retCode == SQL_SUCCESS_WITH_INFO) {
        std::wcout << L"Insert successfully executed.\n";
    }
    else {
        std::wcout << L"Failed to execute INSERT query.\n";
    }
}

void Update(SQLRETURN retCode, SQLHSTMT hStmt, SQLHDBC hDbc,SQLINTEGER id,SQLINTEGER newAge) {
    SQLWCHAR query[] = L"Update SinhVien set age = ? where id = ?";
    retCode = SQLPrepare(hStmt, query, SQL_NTS);
    if (retCode != SQL_SUCCESS && retCode != SQL_SUCCESS_WITH_INFO) {
        std::wcout << L"Error in Update SQL\n";
        return;
    }
    SQLBindParameter(hStmt, 2, SQL_PARAM_INPUT, SQL_C_SLONG, SQL_INTEGER, 0, 0, &id, 0, nullptr);
    SQLBindParameter(hStmt, 1, SQL_PARAM_INPUT, SQL_C_SLONG, SQL_INTEGER, 0, 0, &newAge, 0, nullptr);

    retCode = SQLExecute(hStmt);
    if (retCode == SQL_SUCCESS || retCode == SQL_SUCCESS_WITH_INFO) {
        std::wcout << L"Update successfully executed.\n";
    }
    else {
        std::wcout << L"Failed to execute Update query.\n";
    }
}

void Delete(SQLRETURN retCode, SQLHSTMT hStmt, SQLHDBC hDbc, SQLINTEGER id) {
    SQLWCHAR query[] = L"DELETE FROM SinhVien  where id = ?";
    retCode = SQLPrepare(hStmt, query, SQL_NTS);
    if (retCode != SQL_SUCCESS && retCode != SQL_SUCCESS_WITH_INFO) {
        std::wcout << L"Error in Update SQL\n";
        return;
    }
    SQLBindParameter(hStmt, 1, SQL_PARAM_INPUT, SQL_C_SLONG, SQL_INTEGER, 0, 0, &id, 0, nullptr);

    retCode = SQLExecute(hStmt);
    if (retCode == SQL_SUCCESS || retCode == SQL_SUCCESS_WITH_INFO) {
        std::wcout << L"DELETE successfully executed.\n";
    }
    else {
        std::wcout << L"Failed to execute DELETE query.\n";
    }
}
void printMenu() {
    std::cout << "Menu" << std::endl;
    std::cout << "1. Get Data.\n";
    std::cout << "2. Insert SinhVien\n";
    std::cout << "3. Update Age\n";
    std::cout << "4. Delete SinhVien\n";
    std::cout << "5. Export\n";
    std::wcout << "6. Thoát.\n";
    std::cout << "Select: ";
}
int main() {
    SQLHENV hEnv;       // Environment handle
    SQLHDBC hDbc;       // Connection handle
    SQLHSTMT hStmt;     // Statement handle
    SQLRETURN retCode;  // Return code

    // Allocate environment
    retCode = SQLAllocHandle(SQL_HANDLE_ENV, SQL_NULL_HANDLE, &hEnv);
    checkError(retCode, hEnv, SQL_HANDLE_ENV);

    // Set ODBC version
    retCode = SQLSetEnvAttr(hEnv, SQL_ATTR_ODBC_VERSION, (void*)SQL_OV_ODBC3, 0);
    checkError(retCode, hEnv, SQL_HANDLE_ENV);

    // Allocate connection
    retCode = SQLAllocHandle(SQL_HANDLE_DBC, hEnv, &hDbc);
    checkError(retCode, hDbc, SQL_HANDLE_DBC);

    // Connect to the database
    SQLWCHAR connStr[] = L"DRIVER={ODBC Driver 17 for SQL Server};SERVER=V000389\\SQLEXPRESS;DATABASE=SinhVien;Trusted_Connection=Yes;";  // Replace with your DSN
    retCode = SQLDriverConnect(hDbc, nullptr, connStr, SQL_NTS, nullptr, 0, nullptr, SQL_DRIVER_COMPLETE);
    checkError(retCode, hDbc, SQL_HANDLE_DBC);

    std::wcout << L"Connected to SQL Server!\n";

    // Allocate statement
    retCode = SQLAllocHandle(SQL_HANDLE_STMT, hDbc, &hStmt);
    checkError(retCode, hStmt, SQL_HANDLE_STMT);
    
    while (true) {
        bool flag = true;
        printMenu();
        int select;
        std::cin >> select;
        switch (select) {
            case 1:
                cls;
                GetData(retCode, hStmt, hDbc);
                std::cout << "\n";
                break;
            case 2:
                cls;
                SQLINTEGER id, age;
                SQLWCHAR name[100];
                std::cout << "ID: ";std::cin >> id;
                std::cin.ignore();
                std::cout << "Name: ";std::wcin.getline(name, sizeof(name)/sizeof(SQLWCHAR));
                std::cout << "Age: ";std::cin >> age;
                Insert(retCode, hStmt, hDbc, id, name, age);
                std::cout << "\n";
                break;
            case 3:
                cls;
                std::cout << "ID: ";std::cin >> id;
                std::cout << "Age: ";std::cin >> age;
                Update(retCode, hStmt, hDbc, id, age);
                std::cout << "\n";
                break;
            case 4:
                cls;
                std::cout << "ID: ";std::cin >> id;
                Delete(retCode, hStmt, hDbc, id);
                std::cout << "\n";
                break;
            case 5:
                exportDataToExcel(retCode, hStmt, hDbc, "output.xlsx");
                break;
            case 6:
                flag = false;
                break;
            default:
                std::wcout << "Không hợp lệ!\n";
                break;
        }
        if (!flag) {
            cls;
            std::cout << "Exited\n";
            break;
        }
        pause;
        cls;
    }
    // GET DATA SQL
    //GetData(retCode, hStmt, hDbc);

    // Insert data
    //SQLWCHAR name[] = L"Phuc 3";
    //Insert(retCode, hStmt, hDbc, 3, name, 20);
    //GetData(retCode, hStmt, hDbc);
    
    //Update Data
    //Update(retCode, hStmt, hDbc, 3, 100);
    //GetData(retCode, hStmt, hDbc);

    //Delete
    /*Delete(retCode, hStmt, hDbc, 3);
    GetData(retCode, hStmt, hDbc);*/

    // Cleanup
    SQLFreeHandle(SQL_HANDLE_STMT, hStmt);
    SQLDisconnect(hDbc);
    SQLFreeHandle(SQL_HANDLE_DBC, hDbc);
    SQLFreeHandle(SQL_HANDLE_ENV, hEnv);

    return 0;
}
