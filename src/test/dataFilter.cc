#include <xlnt/xlnt.hpp>
#include <iostream>
#include <iomanip>
#include <string>
#include <regex>

// 辅助函数：比较两个datetime是否在范围内（年月日比较）
bool is_between(const xlnt::datetime& date,
    const xlnt::datetime& start,
    const xlnt::datetime& end) {
    if (date.year < start.year) return false;
    if (date.year > end.year) return false;

    if (date.year == start.year && date.month < start.month) return false;
    if (date.year == end.year && date.month > end.month) return false;

    if (date.year == start.year && date.month == start.month && date.day < start.day) return false;
    if (date.year == end.year && date.month == end.month && date.day > end.day) return false;

    return true;
}

int main() {
    try {
        xlnt::workbook wb;
        wb.load("E:/develop/FilterDataProject/src/test.xlsx");
        xlnt::worksheet ws = wb.active_sheet();

        xlnt::datetime start_date(2023, 1, 1);
        xlnt::datetime end_date(2025, 2, 1);

        char date_column = 'Q';
        char value_column = 'A';
        int start_row = 3;

        int count = 0;
        double total_value = 0.0;

        std::cout << "符合条件的记录：" << std::endl;
        std::cout << "日期\t\t数值" << std::endl;

        for (int row = start_row; row <= ws.highest_row(); ++row) {
            xlnt::cell date_cell = ws.cell(xlnt::cell_reference(date_column, row));

            if (!date_cell.has_value()) {
                std::cout << "cell is empty" << std::endl;
                continue;
            }

            xlnt::datetime cell_date(1900, 1, 1);
            bool is_valid_date = false;

            // 判断是否为日期类型
            if (date_cell.data_type() == xlnt::cell_type::date) {
                cell_date = date_cell.value<xlnt::datetime>();
                is_valid_date = true;
            }
            // 修复：将xlnt::cell_type::string改为xlnt::cell_type::text（适配1.5.0版本）
            else if (date_cell.data_type() == xlnt::cell_type::text) {  // 关键改动处
                std::string date_str = date_cell.value<std::string>();
                std::regex date_regex(R"((\d{4})[/-](\d{1,2})[/-](\d{1,2}))");
                std::smatch match;
                if (std::regex_match(date_str, match, date_regex)) {
                    int year = std::stoi(match[1]);
                    int month = std::stoi(match[2]);
                    int day = std::stoi(match[3]);
                    cell_date = xlnt::datetime(year, month, day);
                    is_valid_date = true;
                }
                else {
                    std::cout << "not date: 字符串无法解析（行号=" << row << ", 值=" << date_str << "）" << std::endl;
                }
            }
            // 判断是否为数字类型
            else if (date_cell.data_type() == xlnt::cell_type::number) {
                double date_num = date_cell.value<double>();
                cell_date = xlnt::datetime::from_number(date_num, xlnt::calendar::windows_1900);
                is_valid_date = true;
            }
            // 其他不支持的类型
            else {
                std::cout << "not date：不支持的类型（行号=" << row << ",类型=" << static_cast<int>(date_cell.data_type()) << "）" << std::endl;
            }

            if (!is_valid_date) {
                std::cout << "not date" << std::endl;
                continue;
            }

            if (is_between(cell_date, start_date, end_date)) {
                xlnt::cell value_cell = ws.cell(xlnt::cell_reference(value_column, row));
                double value = 0.0;

                if (value_cell.has_value() && value_cell.data_type() == xlnt::cell_type::number) {
                    value = value_cell.value<double>();
                }

                count++;
                total_value += value;

                std::cout << static_cast<int>(cell_date.year) << "-"
                    << std::setw(2) << std::setfill('0') << static_cast<int>(cell_date.month) << "-"
                    << std::setw(2) << std::setfill('0') << static_cast<int>(cell_date.day) << "\t"
                    << value << std::endl;
            }
        }

        std::cout << "\n统计结果：" << std::endl;
        std::cout << "符合条件的记录总数：" << count << std::endl;
        std::cout << "数值列总和：" << total_value << std::endl;
        if (count > 0) {
            std::cout << "数值列平均值：" << total_value / count << std::endl;
        }

    }
    catch (const std::exception& e) {
        std::cerr << "发生错误：" << e.what() << std::endl;
        return 1;
    }

    return 0;
}
