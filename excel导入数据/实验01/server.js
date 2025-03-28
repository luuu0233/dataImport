const express = require('express');
const bodyParser = require('body-parser');
const mysql = require('mysql2');
const app = express();
const path = require('path');

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
app.post('/upload', (req, res) => {
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

    // 构建SQL插入语句
    const tableName = 'school2';
    const sql = `INSERT INTO ${tableName} (${columns.join(', ')}) VALUES ?`;

    connection.query(sql, [rows], (error, results) => {
        if (error) {
            console.error('Error inserting data:', error);
            res.status(500).json({
                //success: false,
                status:'error',
                message: 'Error inserting data',
                error: error.message
            });
        } else {
            console.log('Data inserted successfully');
            res.status(200).json({
                //success: true,
                status:'success',
                message: 'Data inserted successfully',
                results: results
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