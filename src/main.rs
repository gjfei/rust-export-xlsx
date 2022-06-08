use regex::Regex;
use simple_excel_writer::*;
use std::fs;
use std::io;

fn main() {
    println!("输入原始文件路径");
    let mut input_file_path = String::new();

    io::stdin()
        .read_line(&mut input_file_path)
        .expect("文件路径解析错误");

    println!("你的文件路径: {}", input_file_path);

    let contents = fs::read_to_string(input_file_path.trim().to_string()).expect("文件读取出错");

    println!("你的文件内容: {}", contents);

    let split_vec: Vec<&str> = contents.split("  sap ").collect();

    let match_vlan = Regex::new(r"lag-(.+?),create").unwrap();
    let match_ip = Regex::new(r"ip (.+),create").unwrap();

    let mut result: Vec<Row> = Vec::new();

    for group_content in split_vec {
        let match_vlan_result = match_vlan.captures_iter(&group_content).collect::<Vec<_>>();

        if match_vlan_result.len() > 0 {
            let vlan = match_vlan_result
                .get(0)
                .unwrap()
                .get(1)
                .unwrap()
                .as_str()
                .trim();
            let match_ip_result = match_ip.captures_iter(&group_content).collect::<Vec<_>>();

            for ip_content in match_ip_result {
                let ip = ip_content.get(1).unwrap().as_str().trim();
                result.push(row![vlan.to_string(), ip.to_string()])
            }
        }
    }

    let mut output_path = String::new();

    println!("输入输出文件路径");

    io::stdin()
        .read_line(&mut output_path)
        .expect("保存路径解析错误");

    output_path = output_path.trim().to_string();

    let mut file_name = input_file_path.split("\\").last().unwrap();

    file_name = file_name.split(".").collect::<Vec<&str>>()[0];

    let output_file_path;

    if output_path.len() == 0 {
        output_file_path = format!("/{}.xlsx", file_name.to_string());
    } else if output_path.contains(".xlsx") {
        output_file_path = output_path;
    } else {
        output_file_path = format!("{}/{}.xlsx", output_path, file_name.to_string());
    }

    println!("输入输出文件路径: {}", output_file_path);

    let mut wb = Workbook::create(&output_file_path);
    let mut sheet = wb.create_sheet("第一页");

    sheet.add_column(Column { width: 30.0 });
    sheet.add_column(Column { width: 30.0 });

    wb.write_sheet(&mut sheet, |sheet_writer| {
        let sw = sheet_writer;
        sw.append_row(row!["vlan", "ip"])?;

        for ele in result {
            sw.append_row(ele)?;
        }
        sw.append_row(row!["", ""])
    })
    .expect("写入excel错误!");

    wb.close().expect("关闭excel错误!");

    println!("写入excel成功!");

    println!("文件路径: {}", output_file_path);

    println!("按回车键退出");

    io::stdin().read_line(&mut String::new()).unwrap();
}
