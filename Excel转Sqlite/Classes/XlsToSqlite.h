
// ����Ȫ
#pragma once
#include <string>

namespace YExcel
{
	class BasicExcelWorksheet;
}

struct sqlite3;

typedef void (*MsgCallBack)(const char* msg);
// @��ע: xls�ļ���������(������������)����ֵ������Ϊ��
class XlsToSqlite
{
public:
	XlsToSqlite(void);
	virtual ~XlsToSqlite(void);
	// @param xlsFile: .xls�ļ�
	// @param sqliteFile: sqlite�ļ�
	// @param sqlPassWord: ����
	// @param skipLines: ����������(����д�ֶ�ע�͵�)
	// @praam callback: ��Ϣ�ص�
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

