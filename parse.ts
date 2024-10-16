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

// 解析金额，将带货币符号和逗号的金额转换为数字
function parseAmount(amountStr: string): number | null {
    if (amountStr.trim() === '--') {
        return null;
    }
    // 提取数字部分，包括可能的负号和小数点
    const numericPart = amountStr.replace(/[^\d\.-]/g, '');
    if (numericPart === '') return null;
    const result = Number(numericPart);
    if (isNaN(result)) return null;
    return result;
}

// 解析每一行文本为交易记录对象
function parseTransactionLine(line: string): TransactionRecord {
    const fields = line.split('\t');  // 使用制表符拆分
    if (fields.length !== 10) {
        throw new Error(`行格式不正确，期望10个字段，实际得到${fields.length}个字段。行内容：${line}`);
    }
    return {
        交易日期: fields[0],
        记账日期: fields[1],
        记账币种: fields[2],
        卡号: fields[3],
        交易类型: fields[4],
        交易描述: fields[5],
        存入金额: parseAmount(fields[6]),
        支出金额: parseAmount(fields[7]),
        交易币种: fields[8],
        交易金额: parseAmount(fields[9])!,
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
    const file_prefix = '2024-9';
    const inputFilePath = path.join(__dirname, 'datas', `${file_prefix}.txt`);  // txt 文件路径
    const outputFilePath = path.join(__dirname, 'outputs', `${file_prefix}.xlsx`);  // 输出 excel 文件路径

    const records = parseTxtFile(inputFilePath);
    exportToExcel(records, outputFilePath);

    console.log(`Excel 文件已创建：${outputFilePath}`);
}

main();
