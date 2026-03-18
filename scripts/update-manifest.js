#!/usr/bin/env node

/**
 * 脚本：更新 manifest.xml 中的 URL
 * 用法: node update-manifest.js <新URL>
 * 示例: node update-manifest.js https://yourcompany.com/word-addin
 */

const fs = require('fs');
const path = require('path');

const newUrl = process.argv[2];

if (!newUrl) {
    console.error('错误: 请提供新的 URL');
    console.log('用法: node update-manifest.js <新URL>');
    console.log('示例: node update-manifest.js https://yourcompany.com/word-addin');
    process.exit(1);
}

// 确保 URL 不以斜杠结尾
const baseUrl = newUrl.replace(/\/$/, '');

const manifestPath = path.join(__dirname, '..', 'manifest.xml');

try {
    let content = fs.readFileSync(manifestPath, 'utf8');
    
    // 替换所有 localhost:3000 的 URL
    const replacements = [
        { from: /https:\/\/localhost:3000/g, to: baseUrl },
        { from: /http:\/\/localhost:3000/g, to: baseUrl.replace('https://', 'http://') }
    ];
    
    replacements.forEach(({ from, to }) => {
        content = content.replace(from, to);
    });
    
    // 写回文件
    fs.writeFileSync(manifestPath, content, 'utf8');
    
    console.log('✓ manifest.xml 已更新');
    console.log(`  所有 URL 已更新为: ${baseUrl}`);
    console.log('\n请检查以下 URL 是否正确:');
    console.log(`  - ${baseUrl}/taskpane.html`);
    console.log(`  - ${baseUrl}/commands.html`);
    console.log(`  - ${baseUrl}/assets/icon-*.png`);
    
} catch (error) {
    console.error('错误:', error.message);
    process.exit(1);
}

