// File: index.js
const babel = require('@babel/register');

// Xử lý tương thích: Bắt chính xác hàm register dù nó bị bọc trong object hay không
const register = babel.default || babel;

register({
    presets: ['@babel/preset-react']
});

require('./server.js');