// 王智泉

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
	//获取缓冲区大小，并申请空间，缓冲区大小事按字节计算的  
	int len = WideCharToMultiByte(CP_ACP, 0, s.c_str(), s.size(), NULL, 0, NULL, NULL);  
	char* buffer = new char[len + 1];  
	//宽字节编码转换成多字节编码  
	WideCharToMultiByte(CP_ACP, 0, s.c_str(), s.size(), buffer, len, NULL, NULL);  
	buffer[len] = '\0';  
	//删除缓冲区并返回值  
	result.append(buffer);  
	delete[] buffer;  
	return result; 
}

std::string s2utf8(const std::string & str) 
{ 
	int nwLen = ::MultiByteToWideChar(CP_ACP, 0, str.c_str(), -1, NULL, 0); 

	wchar_t * pwBuf = new wchar_t[nwLen + 1];//一定要加1，不然会出现尾巴 
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
		// 加载excel
		if (!e.Load(xlsFile))
		{
			LogMsg(std::string("无法打开文件:") + xlsFile);
			return;
		}
	}
	catch( void* e )
	{
		LogMsg(
			"加载Excel文件失败\n"
			"可能原因:\n"
			"	1.请确认指定的Excel路径是否正确!\n"
			"	2.请确保没有其它程序打开指定的Excel文件!\n"
			"	3.请确认所加载的Excel文件为Excel 97-2003格式,即以.xls文件后缀!\n"
			"	4.请确认指定的Excel文件中只有一个工作表，且名称为英文!\n"
			"	5.请确认工作表中的所有字段所对应的列不能有空值!"
			"	6.如果你的表中存在不同编码的文本,请将所有文字复制到文本编辑器，另存为utf-8格式，再重新复制回来保存!\n"
			"	7.其它原因:请将Excel表中的文本复制至文本编辑器并保存,然后再用Excel打开该文本文件，重新保存，确认不存在以上问题后再重新导出!"
			);
		return;
	}
		

	//::DeleteFile(sqlLiteFile);

	// 加载SQLite
	int res = sqlite3_open(sqliteFile, &db);


#ifdef SQLITE_HAS_CODEC


	extern int sqlite3_key(sqlite3 *db, const void *pKey, int nKey);
	if (sqlPassWord)
	{
		sqlite3_key(db, sqlPassWord, strlen(sqlPassWord));
	}
	
#endif

	if( res ){
		LogMsg(std::string("无法打开数据库:") + sqlite3_errmsg(db));
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

		LogMsg("完成!\n");
	}
	//LogMsg("导出成功");
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
	// 得到表名
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

	LogMsg("正在转换表:" + tableName + " ... ");
	
	// 得到行和列的数量
	maxRows = sheet->GetTotalRows();
	maxCols = sheet->GetTotalCols();
	char* errMsg = NULL;
	
	// 删除
	std::string SQL = "DROP TABLE ";
	SQL += tableName;
	int res= sqlite3_exec(db , SQL.c_str() , 0 , 0 , &errMsg);
	if (res != SQLITE_OK)
	{
		//std::cout << "执行SQL 出错." << errMsg << std::endl;
	}
	
	SQL.clear();
	SQL = "CREATE TABLE " + tableName + " (";
	std::string slipt;
	for (size_t c = 0; c < maxCols; ++c)	// 得到字段名
	{
		BasicExcelCell* cell = sheet->Cell(m_skipLines, c);
		
		if(cell->Type() == BasicExcelCell::UNDEFINED || c >= maxCols)
		{
			slipt.empty();
			maxCols = c;		// 表格的宽度只到最后一个非空字段
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
		std::string errorInfo = "执行创建table的SQL 出错.";
		errorInfo += errMsg;
		LogMsg(errorInfo);
		return FALSE;
	}
	else
	{
		
		//std::cout << "创建table的SQL成功执行."<< std::endl;
	}

	return TRUE;
}

// ======================================================================================
int XlsToSqlite::insertValue(YExcel::BasicExcelWorksheet* sheet)
{
	// 得到行和列的数量
	// 得到表名
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

	// 得到键值
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
					cellString = s2utf8(cell->GetString());	// 如果是字符串，将其转换成UTF-8编码
				}
				break;

			case BasicExcelCell::WSTRING:
				{
					cellString = ws2s(cell->GetWString());
					cellString = s2utf8(cellString);	// 如果是字符串，将其转换成UTF-8编码
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
			sprintf(errOut, "导出第'%d'行出错:%s", r + 1, errMsg);
			LogMsg(errOut);
			return FALSE;
		}
	}
	return TRUE;
}
