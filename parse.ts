import * as fs from 'fs';
import * as path from 'path';
import * as xlsx from 'xlsx';

// 定义每一行的数据格式
interface TransactionRecord {
    账户类型: string;
    交易日期: string;
    记账日期: string;
    卡号: string;
    存入金额: number | null;
    支出金额: number | null;
    交易描述: string;
    交易币种: string | null;
    交易金额: number | null;
    汇率: number | null;
}

// 解析金额，将带逗号的金额转换为数字
function parseAmount(amountStr: string): number | null {
    if (amountStr.trim() === '-' || amountStr.trim() === '--') {
        return null;
    }
    // 移除逗号并转换为数字
    const numericPart = amountStr.replace(/,/g, '');
    const result = Number(numericPart);
    if (isNaN(result)) return null;
    return result;
}

// 从交易描述中提取交易币种、交易金额和汇率
function parseDescription(description: string): { description: string, currency: string | null, amount: number | null, exchangeRate: number | null } {
    let currency;
    let amount;
    let exchangeRate;

    // 提取方括号中的内容
    const regex = /\[(.*?)\]/g;
    let match;
    while ((match = regex.exec(description)) !== null) {
        const content = match[1];
        if (content.startsWith('CNY')) {
            // 提取金额
            currency = 'CNY';
            amount = parseAmount(content.replace('CNY', '').trim());
        } else if (content.startsWith('汇率')) {
            // 提取汇率
            exchangeRate = parseFloat(content.replace('汇率', '').trim());
        }
    }

    // 移除方括号及其内容，得到纯交易描述
    const pureDescription = description.replace(/\[.*?\]/g, '').trim();

    return {
        description: pureDescription,
        currency,
        amount,
        exchangeRate,
    };
}

// 解析每一行文本为交易记录对象
function parseTransactionLine(line: string): TransactionRecord {
    const fields = line.split('\t');  // 使用制表符拆分
    if (fields.length !== 7) {
        throw new Error(`行格式不正确，期望7个字段，实际得到${fields.length}个字段。行内容：${line}`);
    }

    const accountType = fields[0].trim();
    const transactionDate = fields[1].trim();
    const postingDate = fields[2].trim();
    const cardNumber = fields[3].trim();
    const depositAmount = parseAmount(fields[4].trim());
    const withdrawalAmount = parseAmount(fields[5].trim());
    const rawDescription = fields[6].trim();

    const { description, currency, amount, exchangeRate } = parseDescription(rawDescription);

    return {
        账户类型: accountType,
        交易日期: transactionDate,
        记账日期: postingDate,
        卡号: cardNumber,
        存入金额: depositAmount,
        支出金额: withdrawalAmount,
        交易描述: description,
        交易币种: currency,
        交易金额: amount,
        汇率: exchangeRate,
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
    const file_prefix = '2024-11';
    const inputFilePath = path.join(__dirname, 'datas', `${file_prefix}.txt`);  // txt 文件路径
    const outputFilePath = path.join(__dirname, 'outputs', `${file_prefix}.xlsx`);  // 输出 Excel 文件路径

    const records = parseTxtFile(inputFilePath);
    exportToExcel(records, outputFilePath);

    console.log(`Excel 文件已创建：${outputFilePath}`);
}

main();
