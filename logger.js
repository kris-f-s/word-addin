// 日志记录器 - 将日志输出到文件
// 注意：使用 AddinLogger 避免与 Office.js 内部的 Logger 类冲突
class AddinLogger {
    constructor() {
        this.logs = [];
        this.startTime = new Date();
        this.maxLogs = 10000; // 最大日志条数
    }
    
    // 格式化日志条目
    formatLog(level, message, data = null) {
        const timestamp = new Date().toISOString();
        const logEntry = {
            timestamp: timestamp,
            level: level,
            message: message,
            data: data
        };
        
        // 添加到日志数组
        this.logs.push(logEntry);
        
        // 限制日志数量
        if (this.logs.length > this.maxLogs) {
            this.logs.shift(); // 移除最旧的日志
        }
        
        // 同时输出到控制台
        const consoleMessage = `[${timestamp}] [${level}] ${message}`;
        if (data) {
            console[level === 'ERROR' ? 'error' : level === 'WARN' ? 'warn' : 'log'](consoleMessage, data);
        } else {
            console[level === 'ERROR' ? 'error' : level === 'WARN' ? 'warn' : 'log'](consoleMessage);
        }
        
        return logEntry;
    }
    
    log(message, data = null) {
        return this.formatLog('INFO', message, data);
    }
    
    error(message, error = null) {
        const errorData = error ? {
            message: error.message,
            stack: error.stack,
            name: error.name
        } : null;
        return this.formatLog('ERROR', message, errorData);
    }
    
    warn(message, data = null) {
        return this.formatLog('WARN', message, data);
    }
    
    // 获取所有日志
    getAllLogs() {
        return this.logs;
    }
    
    // 获取日志文本
    getLogText() {
        let text = `=== Word Add-in 日志 ===\n`;
        text += `开始时间: ${this.startTime.toISOString()}\n`;
        text += `日志总数: ${this.logs.length}\n`;
        text += `生成时间: ${new Date().toISOString()}\n\n`;
        
        this.logs.forEach(log => {
            text += `[${log.timestamp}] [${log.level}] ${log.message}\n`;
            if (log.data) {
                if (typeof log.data === 'object') {
                    text += `  数据: ${JSON.stringify(log.data, null, 2)}\n`;
                } else {
                    text += `  数据: ${log.data}\n`;
                }
            }
        });
        
        return text;
    }
    
    // 下载日志文件
    downloadLog() {
        const logText = this.getLogText();
        const blob = new Blob([logText], { type: 'text/plain;charset=utf-8' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `word-addin-log-${new Date().toISOString().replace(/[:.]/g, '-')}.txt`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
        
        this.log('日志文件已下载');
    }
    
    // 清空日志
    clear() {
        this.logs = [];
        this.startTime = new Date();
        this.log('日志已清空');
    }
}

// 创建全局日志实例
const logger = new AddinLogger();

// 导出到全局
window.logger = logger;

