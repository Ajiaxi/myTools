
// 王智泉
#pragma once
#include <string>

namespace YExcel
{
	class BasicExcelWorksheet;
}

struct sqlite3;

typedef void (*MsgCallBack)(const char* msg);
// @备注: xls文件的所有列(包括跳过的行)的数值都不能为空
class XlsToSqlite
{
public:
	XlsToSqlite(void);
	virtual ~XlsToSqlite(void);
	// @param xlsFile: .xls文件
	// @param sqliteFile: sqlite文件
	// @param sqlPassWord: 密码
	// @param skipLines: 跳过的行数(用来写字段注释等)
	// @praam callback: 信息回调
	void convert(const char* xlsFile, const char* sqliteFile, const char* sqlPassWord, int skipLines = 0, MsgCallBack callback = NULL);

private:

	void parserSheet(YExcel::BasicExcelWorksheet* sheet);

	int createTable(YExcel::BasicExcelWorksheet* sheet);

	int insertValue(YExcel::BasicExcelWorksheet* sheet);

private:

	void LogMsg(const std::string& str);

	sqlite3* db;

	size_t maxRows;
	size_t maxCols;

	int		m_skipLines;

	MsgCallBack m_callback;
};

