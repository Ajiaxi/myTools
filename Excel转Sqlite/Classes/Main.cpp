
#include <string>
#include <iostream>
#include "XlsToSqlite.h"

using namespace std;

void help(char* argv[]) {

	std::string exe = argv[0];
	int pos = exe.find_last_of("\\");
	if (pos != std::string::npos)
	{
		exe = exe.substr(pos + 1, std::string::npos);
	}
	
	std::cout	<< "�÷�:\n" 
				<< "\t" << exe << " src dst skipLines [pwd cryptType]\n"
				<< "\t��:exampe.xls mySqlite.db 123456 0\n"
				<< "\t src \tXls�ļ�·��\n"
				<< "\t dst \tqlite3���ݿ��ļ�·��\n"
				<< "\t skipLines \t����������,������д�ֶε�ע��\n"
				<< "\t pwd \t����\n"
				<< "\t cryptType \t���ܷ�ʽ: Ĭ�� 0 Ŀǰֻ֧��:xxtea ��:0\n"
				<< std::endl;
}


void printMsg(const char* str) {

	printf(str);
}



int main(int argc, char* argv[]) {

	if (argc < 3)
	{
		help(argv);
		return 1;
	}
	std::string src = argv[1];
	std::string dst = argv[2];

	if (argc == 4)
	{
		XlsToSqlite db;
		db.convert(argv[1], argv[2], NULL, atoi(argv[3]), printMsg);
	}
	else if (argc >= 5)
	{
		XlsToSqlite db;
		db.convert(argv[1], argv[2],  argv[4], atoi(argv[3]), printMsg);
	}

	return 0;
}