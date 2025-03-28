const express = require('express');
const bodyParser = require('body-parser');
const mysql = require('mysql2');
const mammoth = require('mammoth');
const fs = require('fs');
const path = require('path');
const app = express();

app.use(bodyParser.json());

// 创建数据库连接
const connection = mysql.createConnection({
    host: 'localhost',
    user: 'root',
    password: 'luxinyue...',
    database: 'school'
});

connection.connect();

// 处理上传的Excel数据
app.post('/upload', async (req, res) => {
    const data = req.body;

    // 假设Excel的第一行是列名
    const columns = data[0];
    const rows = data.slice(1);

    // 获取本地时间并格式化为 MySQL 的 DATETIME 格式
    const getLocalTime = () => {
        const now = new Date();
        const year = now.getFullYear();
        const month = String(now.getMonth() + 1).padStart(2, '0'); // 月份从 0 开始，需要加 1
        const day = String(now.getDate()).padStart(2, '0');
        const hours = String(now.getHours()).padStart(2, '0');
        const minutes = String(now.getMinutes()).padStart(2, '0');
        const seconds = String(now.getSeconds()).padStart(2, '0');
        return `${year}-${month}-${day} ${hours}:${minutes}:${seconds}`;
    };

    const currentTime = getLocalTime(); // 获取本地时间
    columns.push('创建时间'); // 添加时间字段名
    rows.forEach(row => row.push(currentTime)); // 为每一行数据添加当前时间

    // 找到“介绍”列的索引
    const introductionIndex = columns.indexOf('介绍');

    // 遍历每一行数据
    for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        const wordFilePath = row[introductionIndex];

        if (wordFilePath) {
            try {
                // 读取Word文档内容并转换为HTML
                const result = await mammoth.convertToHtml({ path: wordFilePath });
                const htmlContent = result.value;

                // 将HTML内容替换到当前行的“介绍”列
                row[introductionIndex] = htmlContent;
            } catch (error) {
                console.error('Error reading Word file:', error);
                row[introductionIndex] = ''; // 如果读取失败，设置为空字符串
            }
        }
    }

    // 构建SQL插入语句
    const tableName = 'school3'; // 使用 school3 表
    const sql = `INSERT INTO ${tableName} (${columns.join(', ')}) VALUES ?`;

    connection.query(sql, [rows], (error, results) => {
        if (error) {
            console.error('Error inserting data:', error);
            // 返回JSON格式的错误响应
            res.status(500).json({
                status: 'error',
                message: 'Error inserting data into database',
                error: error.message
            });
        } else {
            console.log('Data inserted successfully');
            // 返回JSON格式的成功响应
            res.status(200).json({
                status: 'success',
                message: 'Data inserted successfully',
                insertedRows: results.affectedRows
            });
        }
    });
});

// 设置静态文件目录（假设你的 HTML 文件在 "public" 文件夹中）
app.use(express.static(path.join(__dirname, 'public')));

// 设置根路由
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.listen(3000, () => {
    console.log('Server is running on port 3000');
});