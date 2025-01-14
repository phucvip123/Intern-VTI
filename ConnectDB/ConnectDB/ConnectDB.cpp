#include <iostream>
#include <windows.h>
#include <sql.h>
#include <sqlext.h>

void checkError(SQLRETURN retCode, SQLHANDLE handle, SQLSMALLINT type) {
    if (retCode != SQL_SUCCESS && retCode != SQL_SUCCESS_WITH_INFO) {
        SQLWCHAR sqlState[1024];
        SQLWCHAR message[1024];
        SQLGetDiagRec(type, handle, 1, sqlState, nullptr, message, 1024, nullptr);
        std::cerr << "SQL Error: " << message << "\n";
        exit(-1);
    }
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
    SQLWCHAR connStr[] = L"DRIVER={ODBC Driver 17 for SQL Server};SERVER=PHUC\\SQLEXPRESS;DATABASE=SinhVien;Trusted_Connection=Yes;";  // Replace with your DSN
    retCode = SQLDriverConnect(hDbc, nullptr, connStr, SQL_NTS, nullptr, 0, nullptr, SQL_DRIVER_COMPLETE);
    checkError(retCode, hDbc, SQL_HANDLE_DBC);

    std::wcout << L"Connected to SQL Server!\n";

    // Allocate statement
    retCode = SQLAllocHandle(SQL_HANDLE_STMT, hDbc, &hStmt);
    checkError(retCode, hStmt, SQL_HANDLE_STMT);

    // GET DATA SQL
    GetData(retCode, hStmt, hDbc);

    // Insert data
    SQLWCHAR name[] = L"Phuc 3";
    //Insert(retCode, hStmt, hDbc, 3, name, 20);
    //GetData(retCode, hStmt, hDbc);
    
    //Update Data
    //Update(retCode, hStmt, hDbc, 3, 100);
    //GetData(retCode, hStmt, hDbc);

    //Delete
    Delete(retCode, hStmt, hDbc, 3);
    GetData(retCode, hStmt, hDbc);

    // Cleanup
    SQLFreeHandle(SQL_HANDLE_STMT, hStmt);
    SQLDisconnect(hDbc);
    SQLFreeHandle(SQL_HANDLE_DBC, hDbc);
    SQLFreeHandle(SQL_HANDLE_ENV, hEnv);

    return 0;
}
