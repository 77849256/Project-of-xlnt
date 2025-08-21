#include <iostream>
#include <xlnt/xlnt.hpp>
#include <string>
#include <sstream>
#include <iomanip>
#include <thread>
#include <chrono>
using namespace std;
#include "column_to_letter.h"
#include "isDate.h"
#include "between.h"
#include "print_cell_as_date.h"
#include "str_to_date.h"

//辅助函数：将字符类型数字转化成整型
int ch_to_int(char num) {
	//检查字符是否为数字（0~9）
	if (num >= '0' && num <= '9') {
		//利用ASCII码差值计算，‘0’的ASCII码是48，‘1’是49，以此类推
		return num - '0';
	}
	else {
		//非数字字符返回-1
		return -1;
	}
}

int main(){
	xlnt::workbook wb;	//申请工作簿内存
	string filename;
	cout << "请输入excel文件名：";
	cin >> filename;
	cout << endl;
	//加载已有工作簿至内存
	try {
		wb.load(filename);
		cout << "文件加载成功！" << endl;
	}
	catch (const std::exception& e) {
		std::cerr << "读取文件时发生错误：" << e.what() << std::endl;
		std::this_thread::sleep_for(std::chrono::seconds(3));
		return 1;
	}

	//获取工作表至ws中
	auto ws = wb.active_sheet();

	//获取精确的已使用区域范围并输出（方便调试）
	xlnt::range_reference dim = ws.calculate_dimension();
	if (dim.height() == 0 || dim.width() == 0){
		cout << "工作表为空，无数据可遍历" << endl;
		return 0;
	}
	////获取边界单元格（左上角、右下角）
	//xlnt::cell top_left_cell = ws.cell(dim.top_left());	//获取左上角单元格
	//xlnt::cell bottom_right_cell = ws.cell(dim.bottom_right());//获取右下角单元格
	////从单元格提取行列号
	//xlnt::row_t start_row = top_left_cell.row();	//起始行号
	//xlnt::row_t end_row = bottom_right_cell.row();	//结束行号
	//xlnt::column_t start_col = top_left_cell.column_index();	//起始列索引（如A是1）
	//xlnt::column_t end_col = bottom_right_cell.column_index();	//结束列索引


	string start_cell_str,end_cell_str;//用来保存用户输入的遍历范围string类型（如"A1"--"B2")

	//4、按行列索引遍历（1-based，与Excel一致）
	//for (xlnt::row_t row = start_row; row <= end_row; ++row) {

	//}
	
	//获取单元格
	//auto cell = ws.cell("Q1501");

	int choose = 0;
	//循环操作
	do {
		cout << "请做出你的选择：" << endl;
		cout << "\t查询（请输入1按回车）" << endl;
		cout << "\t退出（请输入2按回车）" << endl;
		cout << "请输入：";
		cin >> choose;
		cout << endl;
		switch (choose) {
		case 1:{
		//让用户输入单元格坐标 确定遍历范围
		cout << "请确定遍历范围：" << endl;
		cout << "起始单元格坐标为：";
		cin >> start_cell_str;
		cout << "终止单元格坐标为：";
		cin >> end_cell_str;
		//将string类型转换为四个行列号
		//先用string构造xlnt::cell类型
		xlnt::cell start_cell = ws[start_cell_str];
		xlnt::cell end_cell = ws[end_cell_str];

		//从单元格提取行列号
		xlnt::row_t start_row = start_cell.row();	//起始行号
		xlnt::row_t end_row = end_cell.row();	//结束行号
		xlnt::column_t start_col = start_cell.column_index();	//起始列索引（如a是1）
		xlnt::column_t end_col = end_cell.column_index();	//结束列索引


		// 调试输出遍历范围
		cout << "表格数据范围：从"
			<< column_to_letter(start_cell) << start_row << "至"
			<< column_to_letter(end_cell) << end_row << endl << endl;

		//让用户输入日期 确定筛选范围
		string startDate;
		cout << "请指定日期筛选范围（格式如2020-01-01或2020/01/01）：" << endl;
		cout << "\t起始日期为：";
		cin >> startDate;
		//用string构造xlnt::date类型 start
		xlnt::date start(1900,1,1);
		try {
			start = str_to_date(startDate);
		}
		catch (const exception& e) {
			cerr << "起始日期格式错误：" << e.what() << endl;
			break;
		}
		string endDate;
		cout << "\t终止日期为：";
		cin >> endDate;
		//用string构造xlnt::date类型 end
		xlnt::date end(1900, 1, 1);
		try {
			end = str_to_date(endDate);
		}catch (const exception& e) {
			cerr << "终止日期格式错误：" << e.what() << endl;
			break;
		}
		cout << "遍历中……" << endl;
		//std::this_thread::sleep_for(std::chrono::seconds(3));

		int count = 0;//统计变量
		//遍历操作工作表ws	(可更改遍历范围）
		//for (auto row : ws.rows(false)) {
		for (xlnt::row_t row = start_row; row <= end_row;++row) {
			//for (auto cell : row) {
			for (xlnt::column_t col = start_col; col <= end_col; ++col) {
				xlnt::cell cell = ws.cell(col, row);
				if (!cell.has_value()) {
					continue;
				}
				try {
					//对 cell 判断操作1 是否为日期
					if (isDate(cell)) {	//是日期格式
						if (between(cell, start, end)) {	//判断该日期是否在指定范围内
							count++;
							cout << "单元格位置" << column_to_letter(cell) << cell.row() << "检测到日期:";
							print_cell_as_date(cell);
							cout << endl;
						}
					}
				}catch (const xlnt::invalid_attribute& e) {
					cerr << "单元格" << column_to_letter(cell) << row << "属性错误：" << e.what() << endl;
					continue;
				}
				//else {	//不是日期格式，跳过当前循环
				//	//cout << "不是日期格式" << endl;
				//	continue;
				//}
			}
		}
		cout << "检测到结果共  " << count << "个" << endl << endl;
		break;
		}
		case 2:{
			cout << "再见，欢迎您的再次使用！" << endl;
			return 0;
			}
		default:
			cout << "无效输入，请重新选择！" << endl;
		}
	} while (1);
	return 0;
}