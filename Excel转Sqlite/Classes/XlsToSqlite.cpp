// ����Ȫ

#include "XlsToSqlite.h"
#include "BasicExcel.hpp"
#include "Windows.h"
#include "assert.h"

#include <vector>
#include <string>

extern "C" {
#include "sqlite3.h"
};

#ifndef SQLITE_HAS_CODEC
#define SQLITE_HAS_CODEC
#endif

std::wstring s2ws(const std::string& s)
{
	int len;
	int slength = (int)s.length() + 1;
	len = MultiByteToWideChar(CP_ACP, 0, s.c_str(), slength, 0, 0); 
	std::wstring r(len, L'\0');
	MultiByteToWideChar(CP_ACP, 0, s.c_str(), slength, &r[0], len);
	return r;
}

std::string ws2s(const std::wstring& s)
{
	string result;  
	//��ȡ��������С��������ռ䣬��������С�°��ֽڼ����  
	int len = WideCharToMultiByte(CP_ACP, 0, s.c_str(), s.size(), NULL, 0, NULL, NULL);  
	char* buffer = new char[len + 1];  
	//���ֽڱ���ת���ɶ��ֽڱ���  
	WideCharToMultiByte(CP_ACP, 0, s.c_str(), s.size(), buffer, len, NULL, NULL);  
	buffer[len] = '\0';  
	//ɾ��������������ֵ  
	result.append(buffer);  
	delete[] buffer;  
	return result; 
}

std::string s2utf8(const std::string & str) 
{ 
	int nwLen = ::MultiByteToWideChar(CP_ACP, 0, str.c_str(), -1, NULL, 0); 

	wchar_t * pwBuf = new wchar_t[nwLen + 1];//һ��Ҫ��1����Ȼ�����β�� 
	ZeroMemory(pwBuf, nwLen * 2 + 2); 

	::MultiByteToWideChar(CP_ACP, 0, str.c_str(), str.length(), pwBuf, nwLen); 

	int nLen = ::WideCharToMultiByte(CP_UTF8, 0, pwBuf, -1, NULL, NULL, NULL, NULL); 

	char * pBuf = new char[nLen + 1]; 
	ZeroMemory(pBuf, nLen + 1); 

	::WideCharToMultiByte(CP_UTF8, 0, pwBuf, nwLen, pBuf, nLen, NULL, NULL); 

	std::string retStr(pBuf); 

	delete []pwBuf; 
	delete []pBuf; 

	pwBuf = NULL; 
	pBuf  = NULL; 

	return retStr; 
} 

using namespace YExcel;

XlsToSqlite::XlsToSqlite(void)
: maxRows(0)
, maxCols(0)
, m_callback(NULL)
{
}


XlsToSqlite::~XlsToSqlite(void)
{
}

void XlsToSqlite::LogMsg(const std::string& str) {
	if (m_callback)
	{
		(*m_callback)(str.c_str());
	}
}

void XlsToSqlite::convert(const char* xlsFile, const char* sqliteFile, const char* sqlPassWord, int skipLines, MsgCallBack callback)
{
	BasicExcel e;
	m_skipLines = skipLines;
	m_callback = callback;
	try
	{
		// ����excel
		if (!e.Load(xlsFile))
		{
			LogMsg(std::string("�޷����ļ�:") + xlsFile);
			return;
		}
	}
	catch( void* e )
	{
		LogMsg(
			"����Excel�ļ�ʧ��\n"
			"����ԭ��:\n"
			"	1.��ȷ��ָ����Excel·���Ƿ���ȷ!\n"
			"	2.��ȷ��û�����������ָ����Excel�ļ�!\n"
			"	3.��ȷ�������ص�Excel�ļ�ΪExcel 97-2003��ʽ,����.xls�ļ���׺!\n"
			"	4.��ȷ��ָ����Excel�ļ���ֻ��һ��������������ΪӢ��!\n"
			"	5.��ȷ�Ϲ������е������ֶ�����Ӧ���в����п�ֵ!"
			"	6.�����ı��д��ڲ�ͬ������ı�,�뽫�������ָ��Ƶ��ı��༭�������Ϊutf-8��ʽ�������¸��ƻ�������!\n"
			"	7.����ԭ��:�뽫Excel���е��ı��������ı��༭��������,Ȼ������Excel�򿪸��ı��ļ������±��棬ȷ�ϲ�������������������µ���!"
			);
		return;
	}
		

	//::DeleteFile(sqlLiteFile);

	// ����SQLite
	int res = sqlite3_open(sqliteFile, &db);


#ifdef SQLITE_HAS_CODEC


	extern int sqlite3_key(sqlite3 *db, const void *pKey, int nKey);
	if (sqlPassWord)
	{
		sqlite3_key(db, sqlPassWord, strlen(sqlPassWord));
	}
	
#endif

	if( res ){
		LogMsg(std::string("�޷������ݿ�:") + sqlite3_errmsg(db));
		sqlite3_close(db);
		return;
	}
	size_t maxSheets = e.GetTotalWorkSheets();
	for (size_t i = 0; i < maxSheets; ++i)
	{
		YExcel::BasicExcelWorksheet* sheet = e.GetWorksheet(i);
		if (sheet)
		{
			this->parserSheet(sheet);
		}

		LogMsg("���!\n");
	}
	//LogMsg("�����ɹ�");
	//sqlite3_close(db);
}

// ======================================================================================
void XlsToSqlite::parserSheet(YExcel::BasicExcelWorksheet* sheet)
{
	if (NULL == sheet)
	{
		return;
	}

	if (this->createTable(sheet))
	{
		this->insertValue(sheet);
	}	
}

// ======================================================================================
int XlsToSqlite::createTable(YExcel::BasicExcelWorksheet* sheet)
{
	if (NULL == sheet)
	{
		return FALSE;
	}
	// �õ�����
	std::string tableName;
	wchar_t* wTableName = sheet->GetUnicodeSheetName();
	if (wTableName)
	{
		tableName = ws2s(wTableName);
	} else
	{
		tableName = sheet->GetAnsiSheetName();
	}

	assert(!tableName.empty());

	LogMsg("����ת����:" + tableName + " ... ");
	
	// �õ��к��е�����
	maxRows = sheet->GetTotalRows();
	maxCols = sheet->GetTotalCols();
	char* errMsg = NULL;
	
	// ɾ��
	std::string SQL = "DROP TABLE ";
	SQL += tableName;
	int res= sqlite3_exec(db , SQL.c_str() , 0 , 0 , &errMsg);
	if (res != SQLITE_OK)
	{
		//std::cout << "ִ��SQL ����." << errMsg << std::endl;
	}
	
	SQL.clear();
	SQL = "CREATE TABLE " + tableName + " (";
	std::string slipt;
	for (size_t c = 0; c < maxCols; ++c)	// �õ��ֶ���
	{
		BasicExcelCell* cell = sheet->Cell(m_skipLines, c);
		
		if(cell->Type() == BasicExcelCell::UNDEFINED || c >= maxCols)
		{
			slipt.empty();
			maxCols = c;		// ���Ŀ��ֻ�����һ���ǿ��ֶ�
			break;
		}
		else
		{
			SQL += slipt;
			slipt = ",";
		}

		const char* asciiCellString = cell->GetString();
		if (asciiCellString)
		{
			SQL += asciiCellString;
			SQL += " varchar(0)";
		} 
		else
		{
			SQL += ws2s(cell->GetWString()) + " varchar(0)";
		}

		
	}
	SQL += ")";

	res = sqlite3_exec(db , SQL.c_str() ,0 ,0, &errMsg);

	if (res != SQLITE_OK)
	{
		std::string errorInfo = "ִ�д���table��SQL ����.";
		errorInfo += errMsg;
		LogMsg(errorInfo);
		return FALSE;
	}
	else
	{
		
		//std::cout << "����table��SQL�ɹ�ִ��."<< std::endl;
	}

	return TRUE;
}

// ======================================================================================
int XlsToSqlite::insertValue(YExcel::BasicExcelWorksheet* sheet)
{
	// �õ��к��е�����
	// �õ�����
	std::string tableName;
	wchar_t* wTableName = sheet->GetUnicodeSheetName();
	if (wTableName)
	{
		tableName = ws2s(wTableName);
	} else
	{
		tableName = sheet->GetAnsiSheetName();
	}
	char* errMsg = NULL;
	assert(maxCols > 0 && !tableName.empty());

	// �õ���ֵ
	std::string cellString;
	char tmpStr[256] = {0};
	for (size_t r= m_skipLines + 1; r<maxRows; ++r)
	{
		std::string SQL = "INSERT INTO " + tableName + " VALUES (";
		for (size_t c = 0; c < maxCols; ++c)
		{
			BasicExcelCell* cell = sheet->Cell(r,c);
			cellString.clear();
			switch (cell->Type())
			{
			case BasicExcelCell::UNDEFINED:
				printf("          ");
				break;

			case BasicExcelCell::INT:
				
				sprintf(tmpStr, "%10d", cell->GetInteger());
				cellString = tmpStr;
				break;

			case BasicExcelCell::DOUBLE:
				sprintf(tmpStr, "%10.6lf", cell->GetDouble());
				cellString = tmpStr;
				break;

			case BasicExcelCell::STRING:
				{
					cellString = s2utf8(cell->GetString());	// ������ַ���������ת����UTF-8����
				}
				break;

			case BasicExcelCell::WSTRING:
				{
					cellString = ws2s(cell->GetWString());
					cellString = s2utf8(cellString);	// ������ַ���������ת����UTF-8����
				}
				break;
			}

			cellString   = c < maxCols - 1 && !cellString.empty() ? "'" + cellString + "'," :  "'" + cellString + "'";
			SQL += cellString;
		}
		SQL += ")";
		int res = sqlite3_exec(db , SQL.c_str() ,0 ,0, &errMsg);

		if (res != SQLITE_OK)
		{
			char errOut[1024] = {0};
			sprintf(errOut, "������'%d'�г���:%s", r + 1, errMsg);
			LogMsg(errOut);
			return FALSE;
		}
	}
	return TRUE;
}
