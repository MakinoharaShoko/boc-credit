import * as fs from 'fs';
import * as path from 'path';
import * as xlsx from 'xlsx';

// 定义每一行的数据格式
interface TransactionRecord {
    交易日期: string;
    记账日期: string;
    记账币种: string;
    卡号: string;
    交易类型: string;
    交易描述: string;
    存入金额: number | null;
    支出金额: number | null;
    交易币种: string;
    交易金额: number;
}

// 解析金额，将带逗号的金额转换为数字
function parseAmount(amountStr: string): number | null {
    const result = Number(amountStr.replaceAll(',', '')); // 移除逗号并转换为数字
    if (isNaN(result)) return 0
    return result
}

// 解析每一行文本为交易记录对象
function parseTransactionLine(line: string): TransactionRecord {
    const fields = line.split(/\s+/);  // 使用正则表达式按空格拆分
    return {
        交易日期: fields[0],
        记账日期: fields[1],
        记账币种: fields[2],
        卡号: fields[3],
        交易类型: fields[4],
        交易描述: fields.slice(5, fields.length - 4).join(' '),  // 描述可能包含空格
        存入金额: parseAmount(fields[fields.length - 4]),
        支出金额: parseAmount(fields[fields.length - 3]),
        交易币种: fields[fields.length - 2],
        交易金额: parseAmount(fields[fields.length - 1])!,
    };
}

// 读取和解析文本文件
function parseTxtFile(filePath: string): TransactionRecord[] {
    const data = fs.readFileSync(filePath, 'utf-8');
    const lines = data.split('\n').filter(line => line.trim());  // 去掉空行
    const records: TransactionRecord[] = lines.map(parseTransactionLine);
    return records;
}

// 将数据写入 Excel 文件
function exportToExcel(records: TransactionRecord[], outputFilePath: string) {
    const worksheet = xlsx.utils.json_to_sheet(records);
    const workbook = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(workbook, worksheet, 'Transactions');

    // 写入文件
    xlsx.writeFile(workbook, outputFilePath);
}

// 主程序
function main() {
    const file_prefix = '2024-9'
    const inputFilePath = path.join(__dirname, 'datas', `${file_prefix}.txt`);  // txt 文件路径
    const outputFilePath = path.join(__dirname, 'outputs', `${file_prefix}.xlsx`);  // 输出 excel 文件路径

    const records = parseTxtFile(inputFilePath);
    exportToExcel(records, outputFilePath);

    console.log(`Excel file has been created: ${outputFilePath}`);
}

main();
