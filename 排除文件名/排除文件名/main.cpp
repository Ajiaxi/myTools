#include <fstream>
#include <istream>
#include <string>
#include <set>
#include <iostream>
using namespace std;

// =========================================================================================================
void help(char* argv[]) {

	std::string exe = argv[0];
	int pos = exe.find_last_of("\\");
	if (pos != std::string::npos)
	{
		exe = exe.substr(pos + 1, std::string::npos);
	}

	std::cout << "功能:匹配源文件中的每一行是否都存在于目标文件中，如果不存在则将该行输出于不匹配文件中"
		<< "用法:\n" 
		<< "\t" << exe << " srcFile targetFile unMatchFile"
		<< "\t如:a.txt b.txt c.txt"
		<< "\t srcFile \t要匹配的源文件\n"
		<< "\t targetFile \t要匹配的目标文件\n"
		<< "\t password \t输出不匹配的\n"
		<< std::endl;
}

// =========================================================================================================
void match(const std::string& srcFile, const std::string& targetFile, const std::string& unMatchFile) {
	
	std::cout << "正在加载..." << std::endl;
	ifstream inUsed;
	inUsed.open(targetFile, ifstream::in | ifstream::binary);
	string temp;
	std::set<std::string> set1;
	while (getline(inUsed, temp))
	{
		set1.insert(temp);
	}
	inUsed.close();
	// ===
	std::cout << "正在匹配..." << std::endl;
	ifstream inHas;
	set<string> set2;
	inHas.open(srcFile, ifstream::in | ifstream::binary);
	while(getline(inHas, temp)) {
		if (set1.find(temp) == set1.end())
		{
			set2.insert(temp);
		}
	}
	inHas.close();
	// ==
	ofstream outdata(unMatchFile);
	for (set<string>::iterator pos = set2.begin(); pos != set2.end(); ++pos)
	{
		outdata << *pos << std::endl;
	}
	outdata.close();
	std::cout << "匹配完成!" << std::endl;
}

// ==========================================================================================================
int main(int argc, char* argv[]) {

	if (argc < 4)
	{
		help(argv);
		return 1;
	}

	match(argv[1], argv[2], argv[3]);
	
	return 0;
}