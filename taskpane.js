/* global Office, logger */

// 全局变量声明（必须在任何使用之前）
let ruleCounter = 0;
let isInitialized = false;
let officeReadyPromise = null;
let officeReadyInfo = null;
let isProcessing = false; // 标记是否有操作正在进行

// 添加错误处理和诊断信息
// 注意：logger 可能在此时还未加载，使用 typeof 检查
if (typeof logger !== 'undefined') {
    logger.log('taskpane.js 开始加载...');
}
console.log('taskpane.js 开始加载...');

// 检查是否有操作正在进行，如果有则显示提示并返回 true
function checkIfProcessing() {
    if (isProcessing) {
        showMessage('别点了别点了，我在烧烤。', 'warning', 3000);
        if (typeof logger !== 'undefined') {
            logger.warn('操作正在进行中，阻止新的操作');
        }
        console.warn('操作正在进行中，阻止新的操作');
        return true;
    }
    return false;
}

// 消息显示函数（替代 alert，因为 Office Add-in 不支持 window.alert）
function showMessage(message, type = 'info', duration = 3000) {
    const messageArea = document.getElementById('message-area');
    if (!messageArea) {
        // 如果消息区域不存在，回退到 console.log
        console.log(`[${type.toUpperCase()}] ${message}`);
        return;
    }
    
    // 设置消息内容和样式
    messageArea.textContent = message;
    messageArea.style.display = 'block';
    
    // 根据类型设置颜色
    switch (type) {
        case 'success':
            messageArea.style.backgroundColor = '#d4edda';
            messageArea.style.color = '#155724';
            messageArea.style.border = '1px solid #c3e6cb';
            break;
        case 'error':
            messageArea.style.backgroundColor = '#f8d7da';
            messageArea.style.color = '#721c24';
            messageArea.style.border = '1px solid #f5c6cb';
            break;
        case 'warning':
            messageArea.style.backgroundColor = '#fff3cd';
            messageArea.style.color = '#856404';
            messageArea.style.border = '1px solid #ffeaa7';
            break;
        default: // info
            messageArea.style.backgroundColor = '#d1ecf1';
            messageArea.style.color = '#0c5460';
            messageArea.style.border = '1px solid #bee5eb';
    }
    
    // 自动隐藏
    if (duration > 0) {
        setTimeout(() => {
            messageArea.style.display = 'none';
        }, duration);
    }
}

// 导出到全局
window.showMessage = showMessage;

// 无论 Office.js 是否加载，都尝试初始化基本功能（用于浏览器测试）
// 这样在浏览器中也能测试界面和基本功能

// 等待 DOM 加载完成后初始化
function initWhenReady() {
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', () => {
            if (typeof logger !== 'undefined') {
                logger.log('✓ DOMContentLoaded 事件触发');
            }
            console.log('✓ DOMContentLoaded 事件触发');
            initializeAddIn();
        });
    } else {
        if (typeof logger !== 'undefined') {
            logger.log('✓ DOM 已就绪，直接初始化');
        }
        console.log('✓ DOM 已就绪，直接初始化');
        initializeAddIn();
    }
}

// 无论 Office.js 是否加载，都确保初始化执行
// 先尝试初始化（不等待 Office.js）
if (document.readyState === 'loading') {
    // DOM 还在加载，等待加载完成
    document.addEventListener('DOMContentLoaded', () => {
        console.log('✓ DOMContentLoaded 事件触发（无 Office.js）');
        initializeAddIn();
    });
} else {
    // DOM 已就绪，立即初始化
    console.log('✓ DOM 已就绪，立即初始化（无 Office.js）');
    // 使用 setTimeout 确保在下一个事件循环中执行，避免与 Office.js 检查冲突
    setTimeout(() => {
        initializeAddIn();
    }, 0);
}

// 检查 Office.js 是否加载
if (typeof Office === 'undefined') {
    console.warn('⚠️ Office.js 未加载（在浏览器中测试时这是正常的）');
    
    // 显示友好的提示信息（但不阻止初始化）
    const showInfo = () => {
        const errorDiv = document.createElement('div');
        errorDiv.style.cssText = 'padding: 15px; color: #8a5000; background: #fff4e5; border: 1px solid #ffaa44; border-radius: 4px; margin: 20px;';
        errorDiv.innerHTML = `
            <strong>ℹ️ 提示：</strong>此页面在浏览器中打开时，Office.js 功能不可用。
            但你可以测试界面和基本功能（添加/删除规则行）。
            <br><br>
            <small>要在 Word 中使用完整功能，请在 Word 中加载插件。</small>
        `;
        
        if (document.body) {
            document.body.insertBefore(errorDiv, document.body.firstChild);
        }
    };
    
    if (document.body) {
        showInfo();
    } else {
        document.addEventListener('DOMContentLoaded', showInfo);
    }
} else {
    console.log('✓ Office.js 已加载');
    
    // 使用 Promise 方式调用 Office.onReady，保存 Promise 以便在其他地方使用
    officeReadyPromise = Office.onReady().then((info) => {
        officeReadyInfo = info;
        console.log('✓ Office.onReady 回调执行');
        console.log('  - Host:', info.host);
        console.log('  - Platform:', info.platform);
        
        if (typeof logger !== 'undefined') {
            logger.log('Office.onReady 在页面加载时完成', {
                host: info.host,
                platform: info.platform
            });
        }
        
        if (info.host === Office.HostType.Word) {
            console.log('✓ 检测到 Word 环境');
            // 在 Word 环境中，如果还没初始化，则初始化
            // 使用 setTimeout 避免重复初始化
            setTimeout(() => {
                const container = document.getElementById('rules-container');
                if (!container || container.children.length === 0) {
                    console.log('在 Word 环境中执行初始化...');
                    initWhenReady();
                } else {
                    console.log('已在浏览器中初始化，跳过 Word 环境初始化');
                }
            }, 100);
        } else {
            console.warn('⚠️ 警告: 不在 Word 环境中，Host:', info.host);
            
            // 显示警告信息
            const warningDiv = document.createElement('div');
            warningDiv.style.cssText = 'padding: 15px; color: #8a5000; background: #fff4e5; border: 1px solid #ffaa44; border-radius: 4px; margin: 20px;';
            warningDiv.innerHTML = `
                <strong>⚠️ 警告：</strong>此插件需要在 Microsoft Word 中运行。
                当前环境：${info.host}。
            `;
            
            if (document.body) {
                document.body.insertBefore(warningDiv, document.body.firstChild);
            }
        }
        
        return info;
    }).catch((error) => {
        console.error('✗ Office.onReady 错误:', error);
        
        if (typeof logger !== 'undefined') {
            logger.error('Office.onReady 在页面加载时失败', error);
        }
        
        const errorDiv = document.createElement('div');
        errorDiv.style.cssText = 'padding: 20px; color: #d13438; background: #fef6f6; border: 1px solid #d13438; border-radius: 4px; margin: 20px;';
        errorDiv.innerHTML = `
            <h2 style="margin-top: 0;">✗ Office.js 初始化失败</h2>
            <p><strong>错误信息：</strong> ${error.message || error}</p>
            <p>请查看浏览器控制台获取详细信息。</p>
        `;
        
        if (document.body) {
            document.body.insertBefore(errorDiv, document.body.firstChild);
        }
        
        throw error;
    });
}

function initializeAddIn() {
    // 防止重复初始化
    if (isInitialized) {
        if (typeof logger !== 'undefined') {
            logger.warn('插件已初始化，跳过重复初始化');
        }
        console.log('⚠️ 插件已初始化，跳过重复初始化');
        return;
    }
    
    if (typeof logger !== 'undefined') {
        logger.log('开始初始化插件...');
    }
    console.log('开始初始化插件...');
    
    try {
        // 检查必要的元素是否存在
        const addRuleBtn = document.getElementById('add-rule-btn');
        const confirmBtn = document.getElementById('confirm-btn');
        const container = document.getElementById('rules-container');
        
        if (!addRuleBtn) {
            throw new Error('找不到"添加规则"按钮 (id: add-rule-btn)');
        }
        if (!confirmBtn) {
            throw new Error('找不到"确认"按钮 (id: confirm-btn)');
        }
        if (!container) {
            throw new Error('找不到规则容器 (id: rules-container)');
        }
        
        // 添加第一行规则（如果还没有）
        if (container.children.length === 0) {
            addRuleRow();
            if (typeof logger !== 'undefined') {
                logger.log('✓ 添加了第一行规则');
            }
            console.log('✓ 添加了第一行规则');
        } else {
            if (typeof logger !== 'undefined') {
                logger.log('✓ 规则已存在，跳过添加');
            }
            console.log('✓ 规则已存在，跳过添加');
        }
        
        // 绑定事件（先移除可能存在的旧监听器，避免重复绑定）
        const newAddHandler = () => {
            if (checkIfProcessing()) {
                return; // 如果正在处理，直接返回
            }
            console.log('点击了"添加规则"按钮');
            addRuleRow();
        };
        
        // 移除旧的事件监听器（如果存在）
        addRuleBtn.replaceWith(addRuleBtn.cloneNode(true));
        const newAddBtn = document.getElementById('add-rule-btn');
        newAddBtn.addEventListener('click', newAddHandler);
        console.log('✓ 绑定了"添加规则"按钮事件');
        
        const newConfirmHandler = () => {
            if (checkIfProcessing()) {
                return; // 如果正在处理，直接返回
            }
            if (typeof logger !== 'undefined') {
                logger.log('点击了"确认"按钮');
            }
            console.log('点击了"确认"按钮');
            applyFormatting();
        };
        
        confirmBtn.replaceWith(confirmBtn.cloneNode(true));
        const newConfirmBtn = document.getElementById('confirm-btn');
        newConfirmBtn.addEventListener('click', newConfirmHandler);
        if (typeof logger !== 'undefined') {
            logger.log('✓ 绑定了"确认"按钮事件');
        }
        console.log('✓ 绑定了"确认"按钮事件');
        
        const newResetHandler = () => {
            if (checkIfProcessing()) {
                return; // 如果正在处理，直接返回
            }
            if (typeof logger !== 'undefined') {
                logger.log('点击了"重置"按钮');
            }
            console.log('点击了"重置"按钮');
            resetFormatting();
        };
        
        const resetBtn = document.getElementById('reset-btn');
        if (resetBtn) {
            resetBtn.replaceWith(resetBtn.cloneNode(true));
            const newResetBtn = document.getElementById('reset-btn');
            newResetBtn.addEventListener('click', newResetHandler);
            if (typeof logger !== 'undefined') {
                logger.log('✓ 绑定了"重置"按钮事件');
            }
            console.log('✓ 绑定了"重置"按钮事件');
        }
        
        // 绑定下载日志按钮
        const downloadLogBtn = document.getElementById('download-log-btn');
        if (downloadLogBtn) {
            downloadLogBtn.addEventListener('click', () => {
                // 检查 logger 是否可用（支持多种检查方式）
                const loggerInstance = typeof logger !== 'undefined' ? logger : (typeof window !== 'undefined' && window.logger ? window.logger : null);
                if (loggerInstance) {
                    loggerInstance.log('点击了"下载日志"按钮');
                    loggerInstance.downloadLog();
                } else {
                    showMessage('日志功能未加载。请刷新页面重试。', 'warning', 5000);
                    console.error('Logger 未找到。检查 window.logger:', typeof window.logger);
                }
            });
            if (typeof logger !== 'undefined' || (typeof window !== 'undefined' && window.logger)) {
                const loggerInstance = typeof logger !== 'undefined' ? logger : window.logger;
                loggerInstance.log('✓ 绑定了"下载日志"按钮事件');
            }
        }
        
        // 监听快捷键 Shift+Alt+O+P
        // Mac 键盘对应：Shift + Option + O，然后按 P
        // 注意：Office Add-in 的快捷键支持有限，需要在任务窗格获得焦点时才能工作
        // 如果需要全局快捷键，建议使用 macOS 的全局快捷键工具（如 Keyboard Maestro）
        let waitingForP = false;
        let timeoutId = null;
        
        document.addEventListener('keydown', (e) => {
            // 检测 Shift + Option (Alt) + O
            // Mac 上 altKey 对应 Option 键
            if (e.shiftKey && e.altKey) {
                const key = e.key.toLowerCase();
                
                if (key === 'o' && !waitingForP) {
                    // 按下 O，等待 P
                    waitingForP = true;
                    e.preventDefault();
                    
                    // 清除之前的超时
                    if (timeoutId) {
                        clearTimeout(timeoutId);
                    }
                    
                    // 设置超时，如果 2 秒内没有按 P，重置状态
                    timeoutId = setTimeout(() => {
                        waitingForP = false;
                        timeoutId = null;
                    }, 2000);
                    
                    // 监听下一个按键
                    const pKeyHandler = (e2) => {
                        const key2 = e2.key.toLowerCase();
                        
                        if (key2 === 'p' && waitingForP) {
                            // 成功触发快捷键
                            e2.preventDefault();
                            waitingForP = false;
                            
                            // 清除超时
                            if (timeoutId) {
                                clearTimeout(timeoutId);
                                timeoutId = null;
                            }
                            
                            // 确保任务窗格获得焦点
                            window.focus();
                            
                            // 滚动到顶部（可选）
                            document.body.scrollTop = 0;
                            document.documentElement.scrollTop = 0;
                            
                            console.log('✓ 快捷键 Shift+Option+O+P 已触发');
                        }
                    };
                    
                    // 使用 once: true 自动移除监听器
                    document.addEventListener('keydown', pKeyHandler, { once: true });
                }
            }
            
            // 如果按了其他键（不是 Shift+Option+P），重置状态
            if (waitingForP && !(e.shiftKey && e.altKey)) {
                const key = e.key.toLowerCase();
                if (key !== 'p') {
                    waitingForP = false;
                    if (timeoutId) {
                        clearTimeout(timeoutId);
                        timeoutId = null;
                    }
                }
            }
        });
        console.log('✓ 快捷键监听器已设置');
        
        // 标记为已初始化
        isInitialized = true;
        if (typeof logger !== 'undefined') {
        logger.log('✓ 插件初始化完成！');
        }
        console.log('✓ 插件初始化完成！');
        
    } catch (error) {
        if (typeof logger !== 'undefined') {
        logger.error('✗ 初始化错误', error);
        }
        console.error('✗ 初始化错误:', error);
        
        const errorDiv = document.createElement('div');
        errorDiv.style.cssText = 'padding: 20px; color: #d13438; background: #fef6f6; border: 1px solid #d13438; border-radius: 4px; margin: 20px;';
        errorDiv.innerHTML = `
            <h2 style="margin-top: 0;">✗ 初始化失败</h2>
            <p><strong>错误：</strong> ${error.message}</p>
            <p>请查看浏览器控制台获取详细信息。</p>
        `;
        
        if (document.body) {
            document.body.insertBefore(errorDiv, document.body.firstChild);
        }
    }
}

function addRuleRow() {
    const container = document.getElementById('rules-container');
    if (!container) {
        console.error('找不到 rules-container 元素');
        return;
    }
    
    const ruleId = `rule-${ruleCounter++}`;
    
    const ruleRow = document.createElement('div');
    ruleRow.className = 'rule-row';
    ruleRow.id = ruleId;
    
    ruleRow.innerHTML = `
        <div class="input-group" style="flex: 4;">
            <label>搜索文本</label>
            <input type="text" class="search-text" placeholder="输入要搜索的文本" />
        </div>
        <div class="color-group">
            <label>字体颜色</label>
            <input type="color" class="font-color" value="#000000" />
        </div>
        <div class="color-group">
            <label>背景颜色</label>
            <input type="color" class="bg-color" value="#ffffff" />
        </div>
        <button class="remove-btn" onclick="removeRuleRow('${ruleId}')">×</button>
    `;
    
    container.appendChild(ruleRow);
    if (typeof logger !== 'undefined') {
        logger.log(`✓ 添加了规则行: ${ruleId}`);
    }
    console.log(`✓ 添加了规则行: ${ruleId}`);
}

// 全局函数，供 HTML 中的 onclick 调用
window.addRuleRow = addRuleRow;

function removeRuleRow(ruleId) {
    const ruleRow = document.getElementById(ruleId);
    if (ruleRow) {
        ruleRow.remove();
        if (typeof logger !== 'undefined') {
            logger.log(`✓ 删除了规则行: ${ruleId}`);
        }
        console.log(`✓ 删除了规则行: ${ruleId}`);
    }
    
    // 如果删除后没有规则了，添加一个默认规则
    const container = document.getElementById('rules-container');
    if (container && container.children.length === 0) {
        addRuleRow();
    }
}

window.removeRuleRow = removeRuleRow;

// 将文本设置为黑字并移除高亮的通用函数（在 Word.run 上下文内部使用）
async function restoreTextColorsInContext(context) {
    if (typeof logger !== 'undefined') {
        logger.log('开始将文本设置为黑字并移除高亮（在上下文中）...');
    }
    console.log('开始将文本设置为黑字并移除高亮（在上下文中）...');
    
    try {
        const body = context.document.body;
        
        // 性能优化：尝试最快速的方法 - 直接设置整个 body 的字体颜色
        // 如果文档中有特殊格式的文本，这个方法可能不够精确，但速度最快
        // 如果用户需要更精确的控制，可以回退到段落级别处理
        if (typeof logger !== 'undefined') {
            logger.log('尝试使用 body 级别快速恢复（最快方法）...');
        }
        
        try {
            // 方法1：直接设置整个 body 的字体颜色（最快，但可能不够精确）
            // 注意：这会影响整个文档，包括可能不想改变的格式
            // 但对于"恢复默认"场景，这通常是可以接受的
            body.font.color = '#000000';
            body.font.highlightColor = null;
            
            // 只同步一次
            await context.sync();
            
            if (typeof logger !== 'undefined') {
                logger.log('✓ 已通过 body 级别快速恢复颜色（最快方法）');
            }
            console.log('✓ 已通过 body 级别快速恢复颜色（最快方法）');
            return; // 成功，直接返回
        } catch (bodyError) {
            // 如果 body 级别设置失败，回退到段落级别处理
            if (typeof logger !== 'undefined') {
                logger.warn('body 级别设置失败，回退到段落级别处理', bodyError);
            }
            console.warn('body 级别设置失败，回退到段落级别处理:', bodyError);
        }
        
        // 方法2：段落级别处理（更精确，但较慢）
        if (typeof logger !== 'undefined') {
            logger.log('使用段落级别处理（精确方法）...');
        }
        
        // 获取所有段落
        const paragraphs = body.paragraphs;
        context.load(paragraphs, 'items');
        await context.sync();
        
        const paragraphItems = paragraphs.items;
        const paragraphCount = paragraphItems ? paragraphItems.length : 0;
        
        if (typeof logger !== 'undefined') {
            logger.log(`找到 ${paragraphCount} 个段落`);
        }
        console.log(`找到 ${paragraphCount} 个段落`);
        
        if (paragraphCount === 0) {
            if (typeof logger !== 'undefined') {
                logger.warn('文档中没有段落');
            }
            console.warn('文档中没有段落');
            return;
        }
        
        // 性能优化：一次性设置所有段落的颜色，然后只同步一次
        // 这样可以大幅减少同步次数，提高性能
        for (let i = 0; i < paragraphCount; i++) {
            const paragraph = paragraphItems[i];
            if (!paragraph) continue;
            
            try {
                // 直接设置整个段落的字体颜色和高亮
                paragraph.font.color = '#000000';  // 黑色字体
                paragraph.font.highlightColor = null;  // 移除高亮（无背景）
            } catch (e) {
                if (typeof logger !== 'undefined') {
                    logger.error(`设置段落 ${i} 的颜色失败`, e);
                }
                console.error(`设置段落 ${i} 的颜色失败:`, e);
            }
        }
        
        // 一次性同步所有更改（关键优化：只同步一次）
        await context.sync();
        
        if (typeof logger !== 'undefined') {
            logger.log(`✓ 已将 ${paragraphCount} 个段落设置为黑字并移除高亮`);
        }
        console.log(`✓ 已将 ${paragraphCount} 个段落设置为黑字并移除高亮`);
    } catch (error) {
        if (typeof logger !== 'undefined') {
            logger.error('恢复文本颜色时出错', error);
        }
        console.error('恢复文本颜色时出错:', error);
        throw error;
    }
}

// 将文本设置为黑字并移除高亮的独立函数（用于重置）
async function restoreTextColors() {
    if (typeof logger !== 'undefined') {
        logger.log('开始将文本设置为黑字并移除高亮...');
    }
    console.log('开始将文本设置为黑字并移除高亮...');
    
    try {
        // 检查 Office.js 是否可用
        if (typeof Office === 'undefined') {
            if (typeof logger !== 'undefined') {
                logger.warn('Office.js 未加载，跳过恢复颜色');
            }
            console.warn('Office.js 未加载，跳过恢复颜色');
            return;
        }
        
        // 等待 Office.onReady 完成（使用与 applyFormatting 相同的逻辑）
        if (typeof Word === 'undefined' || !Office.context) {
            if (typeof logger !== 'undefined') {
                logger.log('等待 Office.onReady 完成（恢复颜色）...', {
                    WordDefined: typeof Word !== 'undefined',
                    OfficeContext: Office.context ? 'defined' : 'undefined'
                });
            }
            console.log('等待 Office.onReady 完成（恢复颜色）...');
            
            try {
                let info;
                
                // 如果已经有保存的 info，直接使用
                if (officeReadyInfo) {
                    if (typeof logger !== 'undefined') {
                        logger.log('使用已保存的 Office.onReady 信息（恢复颜色）');
                    }
                    info = officeReadyInfo;
                } else if (officeReadyPromise) {
                    // 如果有正在进行的 Promise，等待它完成
                    if (typeof logger !== 'undefined') {
                        logger.log('等待已存在的 Office.onReady Promise（恢复颜色）...');
                    }
                    info = await officeReadyPromise;
                } else {
                    // 否则，调用 Office.onReady() 并保存 Promise
                    if (typeof logger !== 'undefined') {
                        logger.log('调用 Office.onReady()（恢复颜色）...');
                    }
                    officeReadyPromise = Office.onReady();
                    info = await officeReadyPromise;
                    officeReadyInfo = info;
                }
                
                if (typeof logger !== 'undefined') {
                    logger.log('Office.onReady 完成（恢复颜色）', { 
                        host: info.host, 
                        platform: info.platform 
                    });
                }
                console.log('✓ Office.onReady 完成（恢复颜色）', info);
                
                // 检查是否在 Word 环境中
                if (info.host !== Office.HostType.Word) {
                    if (typeof logger !== 'undefined') {
                        logger.warn('不在 Word 环境中，跳过恢复颜色');
                    }
                    return;
                }
                
                // 检查 Word 对象是否可用
                if (typeof Word === 'undefined') {
                    if (typeof logger !== 'undefined') {
                        logger.warn('Office.onReady 完成后 Word 对象仍不可用，等待...');
                    }
                    console.warn('⚠️ Office.onReady 完成后 Word 对象仍不可用，等待...');
                    
                    // 等待最多 5 秒，每 100ms 检查一次
                    let waitCount = 0;
                    const maxWait = 50; // 50 * 100ms = 5秒
                    
                    while (typeof Word === 'undefined' && waitCount < maxWait) {
                        await new Promise(resolve => setTimeout(resolve, 100));
                        waitCount++;
                    }
                    
                    if (typeof Word === 'undefined') {
                        if (typeof logger !== 'undefined') {
                            logger.warn('Word 对象不可用，跳过恢复颜色');
                        }
                        console.warn('Word 对象不可用，跳过恢复颜色');
                        return;
                    } else {
                        if (typeof logger !== 'undefined') {
                            logger.log(`✓ Word 对象在等待 ${waitCount * 100}ms 后可用（恢复颜色）`);
                        }
                        console.log(`✓ Word 对象在等待 ${waitCount * 100}ms 后可用（恢复颜色）`);
                    }
                }
            } catch (error) {
                if (typeof logger !== 'undefined') {
                    logger.error('等待 Office.onReady 失败（恢复颜色）', error);
                }
                console.error('✗ 等待 Office.onReady 失败（恢复颜色）:', error);
                throw error;
            }
        }
        
        // 再次确认 Word 对象可用
        if (typeof Word === 'undefined') {
            if (typeof logger !== 'undefined') {
                logger.warn('Word 对象不可用，跳过恢复颜色');
            }
            console.warn('Word 对象不可用，跳过恢复颜色');
            return;
        }
        
        // 执行恢复操作
        if (typeof logger !== 'undefined') {
            logger.log('开始执行 Word.run 恢复文本颜色...');
        }
        console.log('开始执行 Word.run 将文本设置为黑字并移除高亮...');
        
        await Word.run(async (context) => {
            await restoreTextColorsInContext(context);
        });
        
        if (typeof logger !== 'undefined') {
            logger.log('✓ 已将文本设置为黑字并移除高亮');
        }
        console.log('✓ 已将文本设置为黑字并移除高亮');
        
    } catch (error) {
        if (typeof logger !== 'undefined') {
            logger.error('✗ 恢复文本颜色时出错', error);
        }
        console.error('✗ 恢复文本颜色时出错:', error);
        throw error;
    }
}

async function applyFormatting() {
    // 检查是否正在处理
    if (checkIfProcessing()) {
        return;
    }
    
    // 设置处理标志
    isProcessing = true;
    
    if (typeof logger !== 'undefined') {
        logger.log('开始应用格式化...');
    }
    console.log('开始应用格式化...');
    
    try {
        const rules = collectRules();
        if (typeof logger !== 'undefined') {
            logger.log('收集到的规则', rules);
        }
        console.log('收集到的规则:', rules);
        
        if (rules.length === 0) {
            if (typeof logger !== 'undefined') {
                logger.warn('没有规则，提示用户');
            }
            showMessage('请至少添加一条规则', 'warning');
            return;
        }
        
        // 检查 Office.js 是否可用
        if (typeof Office === 'undefined') {
            if (typeof logger !== 'undefined') {
                logger.error('Office.js 未加载');
            }
            showMessage('错误: Office.js 未加载。请在 Word 中使用此插件。', 'error');
            return;
        }
        
        // 等待 Office.onReady 完成，确保 Word 对象可用
        // 在 Office Add-in 中，Word 对象需要在 Office.onReady() 完成后才可用
        // 注意：Office.onReady 可能已经在页面加载时被调用过
        if (typeof Word === 'undefined' || !Office.context) {
            if (typeof logger !== 'undefined') {
                logger.log('检查 Office 状态...', {
                    WordDefined: typeof Word !== 'undefined',
                    OfficeContext: Office.context ? 'defined' : 'undefined',
                    OfficeAvailable: typeof Office !== 'undefined',
                    OfficeReadyInfo: officeReadyInfo ? 'exists' : 'null'
                });
            }
            console.log('检查 Office 状态...');
            console.log('Word 对象:', typeof Word !== 'undefined' ? '已定义' : '未定义');
            console.log('Office.context:', Office.context);
            console.log('Office 对象:', typeof Office !== 'undefined' ? '已定义' : '未定义');
            console.log('officeReadyInfo:', officeReadyInfo);
            
            try {
                let info;
                
                // 如果已经有保存的 info，直接使用
                if (officeReadyInfo) {
                    if (typeof logger !== 'undefined') {
                        logger.log('使用已保存的 Office.onReady 信息');
                    }
                    console.log('✓ 使用已保存的 Office.onReady 信息');
                    info = officeReadyInfo;
                } else if (officeReadyPromise) {
                    // 如果有正在进行的 Promise，等待它完成
                    if (typeof logger !== 'undefined') {
                        logger.log('等待已存在的 Office.onReady Promise...');
                    }
                    console.log('等待已存在的 Office.onReady Promise...');
                    info = await officeReadyPromise;
                } else {
                    // 否则，调用 Office.onReady() 并保存 Promise
                    if (typeof logger !== 'undefined') {
                        logger.log('调用 Office.onReady()...');
                    }
                    console.log('调用 Office.onReady()...');
                    officeReadyPromise = Office.onReady();
                    info = await officeReadyPromise;
                    officeReadyInfo = info;
                }
                
                if (typeof logger !== 'undefined') {
                    logger.log('Office.onReady 完成', { 
                        host: info.host, 
                        platform: info.platform 
                    });
                }
                console.log('✓ Office.onReady 完成', info);
                
                // 检查是否在 Word 环境中
                if (info.host !== Office.HostType.Word) {
                    const errorMsg = `此插件只能在 Microsoft Word 中使用。当前环境: ${info.host}`;
                    if (typeof logger !== 'undefined') {
                        logger.error('不在 Word 环境中', { host: info.host });
                    }
                    showMessage('错误: ' + errorMsg, 'error', 5000);
                    return;
                }
                
                // 检查 Word 对象是否可用
                // 在 Microsoft Word for Mac 中，Word 对象可能需要额外的时间加载
                if (typeof Word === 'undefined') {
                    if (typeof logger !== 'undefined') {
                        logger.warn('Office.onReady 完成后 Word 对象仍不可用，等待...');
                    }
                    console.warn('⚠️ Office.onReady 完成后 Word 对象仍不可用，等待...');
                    
                    // 等待最多 5 秒，每 100ms 检查一次
                    let waitCount = 0;
                    const maxWait = 50; // 50 * 100ms = 5秒
                    
                    while (typeof Word === 'undefined' && waitCount < maxWait) {
                        await new Promise(resolve => setTimeout(resolve, 100));
                        waitCount++;
                        if (waitCount % 10 === 0) {
                            // 每 1 秒记录一次日志
                            if (typeof logger !== 'undefined') {
                                logger.log(`等待 Word 对象... (${waitCount * 100}ms)`);
                            }
                            console.log(`等待 Word 对象... (${waitCount * 100}ms)`);
                        }
                    }
                    
                    if (typeof Word === 'undefined') {
                        if (typeof logger !== 'undefined') {
                            logger.error('等待后 Word 对象仍不可用', {
                                waitCount: waitCount,
                                host: info.host,
                                OfficeContext: Office.context ? 'defined' : 'undefined'
                            });
                        }
                        console.error('✗ 等待后 Word 对象仍不可用');
                        console.error('Office.context:', Office.context);
                        showMessage('错误: Word API 未加载。请刷新插件页面或重新加载 Word。', 'error', 5000);
                        return;
                    } else {
                        if (typeof logger !== 'undefined') {
                            logger.log(`✓ Word 对象在等待 ${waitCount * 100}ms 后可用`);
                        }
                        console.log(`✓ Word 对象在等待 ${waitCount * 100}ms 后可用`);
                    }
                } else {
                    if (typeof logger !== 'undefined') {
                        logger.log('✓ Word 对象在 Office.onReady 后立即可用');
                    }
                    console.log('✓ Word 对象在 Office.onReady 后立即可用');
                }
            } catch (error) {
                if (typeof logger !== 'undefined') {
                    logger.error('等待 Office.onReady 失败', error);
                }
                console.error('✗ 等待 Office.onReady 失败:', error);
                showMessage('错误: ' + (error.message || '无法初始化 Word API。请确保在 Microsoft Word 中使用此插件。'), 'error', 5000);
                return;
            }
        }
        
        // 再次确认 Word 对象可用
        if (typeof Word === 'undefined') {
            if (typeof logger !== 'undefined') {
                logger.error('Word 对象仍未定义');
            }
            showMessage('错误: Word API 未加载。请刷新插件页面重试。', 'error', 5000);
            return;
        }
        
        // 确认在 Word 环境中
        if (Office.context && Office.context.host !== Office.HostType.Word) {
            if (typeof logger !== 'undefined') {
                logger.error('不在 Word 环境中', { host: Office.context.host });
            }
            showMessage('错误: 此插件只能在 Microsoft Word 中使用。', 'error');
            return;
        }
        
        if (typeof logger !== 'undefined') {
            logger.log('✓ Word 对象可用，环境检查通过');
        }
        console.log('✓ Word 对象可用，环境检查通过');
        
        if (typeof logger !== 'undefined') {
            logger.log('Office.js 和 Word 对象可用，开始执行 Word.run');
        }
        console.log('✓ Office.js 和 Word 对象可用，开始执行 Word.run');
        
        // 显示加载状态
        const confirmBtn = document.getElementById('confirm-btn');
        const originalText = confirmBtn.querySelector('.ms-Button-label').textContent;
        confirmBtn.querySelector('.ms-Button-label').textContent = '处理中...';
        confirmBtn.disabled = true;
        
        // 尝试执行 Word.run，如果失败会抛出错误
        await Word.run(async (context) => {
            if (typeof logger !== 'undefined') {
                logger.log('✓ 成功进入 Word.run 上下文');
            }
            console.log('✓ 成功进入 Word.run 上下文');
            const body = context.document.body;
            
            // 第一步：恢复全文的原始颜色和背景颜色
            if (typeof logger !== 'undefined') {
            logger.log('第一步：恢复全文原始颜色...');
            }
            console.log('第一步：恢复全文原始颜色...');
            
            try {
                // 性能优化：对于大文档，使用段落级别的处理
                // 这样可以大幅减少需要处理的范围数量，提高性能
                if (typeof logger !== 'undefined') {
                    logger.log('使用段落级别处理，优化大文档性能...');
                }
                
                // 获取所有段落
                const paragraphs = body.paragraphs;
                context.load(paragraphs, 'items');
                await context.sync();
                
                const paragraphItems = paragraphs.items;
                const paragraphCount = paragraphItems ? paragraphItems.length : 0;
                
                if (typeof logger !== 'undefined') {
                    logger.log(`找到 ${paragraphCount} 个段落用于恢复颜色`);
                }
                console.log(`找到 ${paragraphCount} 个段落用于恢复颜色`);
                
                if (paragraphCount === 0) {
                    if (typeof logger !== 'undefined') {
                        logger.warn('文档中没有段落');
                    }
                    console.warn('文档中没有段落');
                    // 继续执行，不返回，因为可能只是没有段落但需要应用规则
                } else {
                    // 性能优化：尝试使用 body 级别快速恢复（最快方法）
                    if (typeof logger !== 'undefined') {
                        logger.log('尝试使用 body 级别快速恢复（最快方法）...');
                    }
                    
                    try {
                        // 方法1：直接设置整个 body 的字体颜色（最快）
                        body.font.color = '#000000';
                        body.font.highlightColor = null;
                        await context.sync();
                        
                        if (typeof logger !== 'undefined') {
                            logger.log('✓ 已通过 body 级别快速恢复颜色（最快方法）');
                        }
                        console.log('✓ 已通过 body 级别快速恢复颜色（最快方法）');
                    } catch (bodyError) {
                        // 如果 body 级别设置失败，使用段落级别处理
                        if (typeof logger !== 'undefined') {
                            logger.warn('body 级别设置失败，使用段落级别处理', bodyError);
                            logger.log(`开始恢复 ${paragraphCount} 个段落的颜色（精确方法）...`);
                        }
                        console.warn('body 级别设置失败，使用段落级别处理:', bodyError);
                        console.log(`开始恢复 ${paragraphCount} 个段落的颜色（精确方法）...`);
                        
                        // 方法2：段落级别处理（更精确，但较慢）
                        // 性能优化：一次性设置所有段落的颜色，然后只同步一次
                        for (let i = 0; i < paragraphCount; i++) {
                            const paragraph = paragraphItems[i];
                            if (!paragraph) continue;
                            
                            try {
                                // 恢复为默认颜色：黑色字体，移除高亮
                                paragraph.font.color = '#000000';  // 黑色字体
                                paragraph.font.highlightColor = null;  // 移除高亮（恢复为无背景）
                            } catch (e) {
                                if (typeof logger !== 'undefined') {
                                    logger.error(`恢复段落 ${i} 的颜色失败`, e);
                                }
                                console.error(`恢复段落 ${i} 的颜色失败:`, e);
                            }
                        }
                        
                        // 一次性同步所有更改（关键优化：只同步一次）
                        await context.sync();
                        
                        if (typeof logger !== 'undefined') {
                            logger.log(`✓ 已恢复 ${paragraphCount} 个段落的颜色`);
                        }
                        console.log(`✓ 已恢复 ${paragraphCount} 个段落的颜色`);
                    }
                }
            } catch (error) {
                if (typeof logger !== 'undefined') {
                    logger.error('恢复颜色时出错', error);
                }
                console.error('恢复颜色时出错:', error);
                // 不抛出错误，继续执行应用规则的步骤
                if (typeof logger !== 'undefined') {
                    logger.warn('恢复颜色失败，但继续执行应用规则...');
                }
                console.warn('恢复颜色失败，但继续执行应用规则...');
            }
            
            // 第二步：按照规则应用新的颜色格式
            if (typeof logger !== 'undefined') {
                logger.log('第二步：应用新规则...');
            }
            console.log('第二步：应用新规则...');
            
            // 对每个规则进行处理
            for (let ruleIndex = 0; ruleIndex < rules.length; ruleIndex++) {
                const rule = rules[ruleIndex];
                
                if (!rule.searchText || rule.searchText.trim() === '') {
                    if (typeof logger !== 'undefined') {
                    logger.warn(`跳过空规则 #${ruleIndex + 1}`);
                    }
                    console.log(`跳过空规则`);
                    continue;
                }
                
                if (typeof logger !== 'undefined') {
                logger.log(`处理规则 #${ruleIndex + 1}: 搜索 "${rule.searchText}", 字体颜色: ${rule.fontColor}, 背景颜色: ${rule.bgColor}`);
                }
                console.log(`处理规则: 搜索 "${rule.searchText}", 字体颜色: ${rule.fontColor}, 背景颜色: ${rule.bgColor}`);
                
                try {
                    // 搜索所有匹配的文本
                    const searchResults = body.search(rule.searchText, {
                        matchCase: false,
                        matchWholeWord: false,
                        matchWildcards: false
                    });
                    
                    // 加载搜索结果
                    context.load(searchResults, 'items');
                    await context.sync();
                    
                    // 在 sync 之后，安全访问 items
                    const items = searchResults.items;
                    const itemCount = items ? items.length : 0;
                    
                    if (typeof logger !== 'undefined') {
                        logger.log(`找到 ${itemCount} 个匹配结果`);
                    }
                    console.log(`找到 ${itemCount} 个匹配结果`);
                    
                    if (itemCount === 0) {
                        if (typeof logger !== 'undefined') {
                            logger.warn(`未找到匹配 "${rule.searchText}" 的文本`);
                        }
                        console.log(`警告: 未找到匹配 "${rule.searchText}" 的文本`);
                        continue;
                    }
                    
                    // 对每个匹配结果应用格式
                    let successCount = 0;
                    let failCount = 0;
                    
                    // 性能优化：在循环外预处理颜色格式，避免重复检查
                    // 确保颜色格式正确（Word.js 接受十六进制颜色字符串）
                    let fontColor = rule.fontColor;
                    let bgColor = rule.bgColor;
                    
                    // 确保颜色值以 # 开头（颜色选择器通常已经包含 #）
                    if (fontColor && !fontColor.startsWith('#')) {
                        fontColor = '#' + fontColor;
                    }
                    if (bgColor && !bgColor.startsWith('#')) {
                        bgColor = '#' + bgColor;
                    }
                    
                    // 判断是否需要移除背景高亮（在循环外计算一次）
                    const shouldRemoveHighlight = !bgColor || bgColor === '#ffffff' || bgColor === '#FFFFFF' || bgColor === '#fff' || bgColor === '#FFF' || bgColor === '#000000' || bgColor === '#000';
                    
                    // 使用 for 循环而不是 forEach，避免在 Word.js 上下文中的问题
                    // 性能优化：先设置所有匹配项的格式，然后只同步一次
                    for (let index = 0; index < itemCount; index++) {
                        const range = items[index];
                        try {
                            // 设置字体颜色
                            range.font.color = fontColor;
                            if (typeof logger !== 'undefined') {
                                logger.log(`  匹配 #${index + 1}: 设置字体颜色 ${fontColor}`);
                            }
                            
                            // 设置背景颜色（高亮颜色）
                            // 如果背景颜色是白色或透明，设置为 null 移除高亮
                            if (shouldRemoveHighlight) {
                                range.font.highlightColor = null;
                                if (typeof logger !== 'undefined') {
                                    logger.log(`  匹配 #${index + 1}: 移除背景高亮`);
                                }
                            } else {
                                range.font.highlightColor = bgColor;
                                if (typeof logger !== 'undefined') {
                                    logger.log(`  匹配 #${index + 1}: 设置背景颜色 ${bgColor}`);
                                }
                            }
                            
                            successCount++;
                        } catch (e) {
                            failCount++;
                            if (typeof logger !== 'undefined') {
                                logger.error(`  匹配 #${index + 1}: 设置颜色失败`, e);
                            }
                            console.error(`  ✗ 设置颜色失败: ${e.message}`);
                        }
                    }
                    
                    // 性能优化：每个规则处理完后只同步一次，而不是每个匹配项都同步
                    await context.sync();
                    if (typeof logger !== 'undefined') {
                    logger.log(`✓ 规则 "${rule.searchText}" 处理完成 (成功: ${successCount}, 失败: ${failCount})`);
                    }
                    console.log(`  ✓ 规则 "${rule.searchText}" 处理完成`);
                } catch (error) {
                    if (typeof logger !== 'undefined') {
                    logger.error(`处理规则 "${rule.searchText}" 时出错`, error);
                    }
                    console.error(`处理规则时出错:`, error);
                    // 继续处理下一个规则
                }
            }
            
            if (typeof logger !== 'undefined') {
            logger.log('✓ 已应用所有规则');
            }
            console.log('✓ 已应用所有规则');
        });
        
        // 恢复按钮状态
        confirmBtn.querySelector('.ms-Button-label').textContent = originalText;
        confirmBtn.disabled = false;
        
        // 显示成功消息
        showMessage('格式应用成功！', 'success');
        if (typeof logger !== 'undefined') {
            logger.log('✓ 格式化应用成功');
        }
        console.log('✓ 格式化应用成功');
        
        // 清除处理标志
        isProcessing = false;
        
    } catch (error) {
        if (typeof logger !== 'undefined') {
            logger.error('✗ 应用格式化时出错', error);
        }
        console.error('✗ 应用格式化时出错:', error);
        console.error('错误详情:', error);
        
        const confirmBtn = document.getElementById('confirm-btn');
        if (confirmBtn) {
            confirmBtn.querySelector('.ms-Button-label').textContent = '确认';
            confirmBtn.disabled = false;
        }
        
        // 清除处理标志
        isProcessing = false;
        
        const errorMsg = error.message || String(error);
        if (typeof logger !== 'undefined') {
            logger.error('错误详情', { 
                message: errorMsg, 
                stack: error.stack,
                name: error.name,
                error: error
            });
        }
        
        // 提供更友好的错误提示
        let userMessage = '错误: ' + errorMsg;
        if (errorMsg.includes('Word') || errorMsg.includes('undefined')) {
            userMessage = '错误: Word API 未正确加载。请尝试刷新插件或重新加载 Word。';
        }
        
        showMessage(userMessage, 'error', 5000);
    }
}

// 重置功能：恢复文本颜色（保留用户输入的规则）
async function resetFormatting() {
    // 检查是否正在处理
    if (checkIfProcessing()) {
        return;
    }
    
    // 设置处理标志
    isProcessing = true;
    
    if (typeof logger !== 'undefined') {
    logger.log('开始重置...');
    }
    console.log('开始重置...');
    
    try {
        // 显示加载状态
        const resetBtn = document.getElementById('reset-btn');
        if (resetBtn) {
            const originalText = resetBtn.querySelector('.ms-Button-label').textContent;
            resetBtn.querySelector('.ms-Button-label').textContent = '处理中...';
            resetBtn.disabled = true;
            
            // 将文本设置为黑字并移除高亮（保留用户输入的规则）
            if (typeof logger !== 'undefined') {
            logger.log('将文本设置为黑字并移除高亮...');
            }
            await restoreTextColors();
            
            // 恢复按钮状态
            resetBtn.querySelector('.ms-Button-label').textContent = originalText;
            resetBtn.disabled = false;
            
            showMessage('重置成功！文本已设置为黑字并移除高亮，规则已保留。', 'success');
            if (typeof logger !== 'undefined') {
                logger.log('✓ 重置完成（规则已保留）');
            }
            console.log('✓ 重置完成（规则已保留）');
            
            // 清除处理标志
            isProcessing = false;
        } else {
            if (typeof logger !== 'undefined') {
            logger.error('找不到 reset-btn 按钮');
            }
        }
        
    } catch (error) {
        if (typeof logger !== 'undefined') {
        logger.error('✗ 重置时出错', error);
        }
        console.error('✗ 重置时出错:', error);
        
        const resetBtn = document.getElementById('reset-btn');
        if (resetBtn) {
            resetBtn.querySelector('.ms-Button-label').textContent = '重置';
            resetBtn.disabled = false;
        }
        
        // 清除处理标志
        isProcessing = false;
        
        const errorMsg = error.message || String(error);
        if (typeof logger !== 'undefined') {
            logger.error('重置错误详情', { message: errorMsg, stack: error.stack });
        }
        showMessage('错误: ' + errorMsg, 'error', 5000);
    }
}

function collectRules() {
    const rules = [];
    const ruleRows = document.querySelectorAll('.rule-row');
    
    ruleRows.forEach((row) => {
        const searchTextInput = row.querySelector('.search-text');
        const fontColorInput = row.querySelector('.font-color');
        const bgColorInput = row.querySelector('.bg-color');
        
        if (searchTextInput && fontColorInput && bgColorInput) {
            const searchText = searchTextInput.value.trim();
            const fontColor = fontColorInput.value;
            const bgColor = bgColorInput.value;
            
            if (searchText) {
                rules.push({
                    searchText: searchText,
                    fontColor: fontColor,
                    bgColor: bgColor
                });
            }
        }
    });
    
    return rules;
}
