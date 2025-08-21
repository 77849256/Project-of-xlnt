#include <xlnt/xlnt.hpp>
#include <iostream>
#include <chrono>
#include <thread>
//#include<string>
using namespace std;

int main() {
	//创建一个工作簿
	string filename;
	cout << "请输入excel文件名：";
	cin >> filename;
	cout << endl;
	xlnt::workbook wb;
	try {
		wb.load(filename);
	}
	catch (const std::exception& e) {
		std::cerr << "读取文件时发生错误：" << e.what() << std::endl;
		std::this_thread::sleep_for(std::chrono::seconds(315));
		return 1;
	}
	//获取第一个工作表
	auto ws = wb.active_sheet();
	// 指定单元格
	string temp;
	cout << "请输入单元格的坐标，将会判断他的类型是否为字符串或公式，并输出字符串或公式，请输入：";
	cin >> temp;
	cout << endl;
	auto cell = ws.cell(temp);
	if (cell.has_value()) {  // 检查是否有值
		//判断类型（1.5.0用cell_type())
		if (cell.data_type() == xlnt::cell_type::shared_string) {
			//直接用cell.as_string()获取值（无需cell_value)
			string str_val = cell.to_string();
			cout << "A2的值（字符串）：" << str_val << endl << endl;
		}
		else {
			cout << "不是字符串" << endl << endl;
		}
	}
	else if (cell.has_formula()) {
		cout << "A2包含公式：" << ws.cell(temp).formula() << endl;
	}
	else {
		cout << "A2没有值和公式" << endl << endl;
	}
	std::this_thread::sleep_for(std::chrono::seconds(315));
	cout << "程序休眠结束" << endl;
	return 0;
}