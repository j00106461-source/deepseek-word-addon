// Office 加载项初始化
Office.onReady((info) => {
    if (info.host === Office.HostType.Word || info.host === Office.HostType.Excel) {
        console.log('DeepSeek 写作助手已加载');
        // 恢复之前保存的设置
        const savedModel = localStorage.getItem('deepseekModel') || 'flash';
        selectModel(savedModel);
    }
});

// 当前选中的模型
let currentModel = 'flash';

// 模型配置
const models = {
    flash: {
        name: 'deepseek-v4-flash',
        displayName: 'DeepSeek V4 Flash',
        cost: '最便宜（推荐）'
    },
    pro: {
        name: 'deepseek-v4-pro',
        displayName: 'DeepSeek V4 Pro',
        cost: '更强但稍贵'
    }
};

// 选择模型
function selectModel(model) {
    currentModel = model;
    localStorage.setItem('deepseekModel', model);
    
    // 更新按钮状态
    document.querySelectorAll('.model-btn').forEach(btn => {
        btn.classList.remove('active');
    });
    event.target.closest('.model-btn').classList.add('active');
    
    // 更新提示信息
    const hints = {
        flash: 'Flash: $0.14/$0.28 per 1M tokens | 完全够用文字任务',
        pro: 'Pro: $0.435/$0.87 per 1M tokens (当前75折) | 更强的推理能力'
    };
    document.getElementById('modelHint').textContent = hints[model];
}

// 保存设置
function saveSettings() {
    const apiKey = document.getElementById('apiKey').value;
    if (!apiKey) {
        showResult('❌ 请输入 API Key', true);
        return;
    }
    if (!apiKey.startsWith('sk-')) {
        showResult('⚠️ API Key 应该以 sk- 开头', true);
        return;
    }
    localStorage.setItem('deepseekApiKey', apiKey);
    showResult('✓ 设置已保存！选定 ' + models[currentModel].displayName, false);
}

// 获取 API Key
function getApiKey() {
    const apiKey = localStorage.getItem('deepseekApiKey');
    if (!apiKey) {
        showResult('❌ 未配置 API Key\n\n请在"设置"中添加你的 DeepSeek API Key\n\n获取步骤:\n1. 访问 platform.deepseek.com/api_keys\n2. 复制你的 API Key\n3. 粘贴到上方设置框\n4. 点击"保存设置"', true);
        return null;
    }
    return apiKey;
}

// 获取选中的文本
async function getSelectedText() {
    try {
        return await Word.run(async (context) => {
            const selection = context.document.getSelection();
            selection.load('text');
            await context.sync();
            return selection.text;
        });
    } catch (error) {
        return null;
    }
}

// 调用 DeepSeek API
async function callDeepSeekAPI(text, prompt, useThinking = false) {
    const apiKey = getApiKey();
    if (!apiKey) return null;
    
    showResult('⏳ 正在分析...', false, true);
    
    try {
        const requestBody = {
            model: models[currentModel].name,
            messages: [
                {
                    role: 'user',
                    content: `${prompt}\n\n---\n\n文本内容：\n${text}`
                }
            ],
            max_tokens: 1024,
            temperature: 0.7
        };
        
        // 如果选择使用思考模式（仅 V4 Pro 支持，且会增加成本）
        if (useThinking && currentModel === 'pro') {
            requestBody.thinking = {
                type: 'enabled',
                budget_tokens: 1024
            };
        }
        
        const response = await fetch('https://api.deepseek.com/chat/completions', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${apiKey}`
            },
            body: JSON.stringify(requestBody)
        });
        
        if (!response.ok) {
            const error = await response.json();
            const errorMsg = error.error?.message || '请求失败';
            
            // 处理常见错误
            if (errorMsg.includes('authentication') || errorMsg.includes('Invalid')) {
                throw new Error('❌ API Key 无效，请检查是否正确');
            } else if (errorMsg.includes('quota')) {
                throw new Error('❌ API 配额已用尽，请充值或等待重置');
            } else if (errorMsg.includes('overloaded')) {
                throw new Error('❌ 服务器繁忙，请稍后重试');
            } else {
                throw new Error(`❌ 错误: ${errorMsg}`);
            }
        }
        
        const data = await response.json();
        
        // 提取返回内容
        const content = data.choices[0].message.content;
        
        // 计算成本（仅用于显示）
        const inputTokens = data.usage.prompt_tokens;
        const outputTokens = data.usage.completion_tokens;
        const costFlash = (inputTokens * 0.14 + outputTokens * 0.28) / 1000000;
        const costPro = (inputTokens * 0.435 + outputTokens * 0.87) / 1000000;
        const estimatedCost = currentModel === 'flash' ? costFlash : costPro;
        
        // 格式化结果
        let resultText = `${content}\n\n---\n💰 估计成本: $${estimatedCost.toFixed(6)} (${inputTokens + outputTokens} tokens)`;
        
        return resultText;
    } catch (error) {
        showResult(`❌ ${error.message}`, true);
        return null;
    }
}

// 显示结果
function showResult(text, isError = false, isLoading = false) {
    const resultDiv = document.getElementById('result');
    if (!text) {
        resultDiv.innerHTML = '';
        return;
    }
    
    let className = 'result';
    if (isError) className += ' error';
    if (isLoading) className += ' loading';
    
    resultDiv.className = className;
    resultDiv.innerHTML = text;
}

// ============ 各个功能按钮 ============

// 语法检查
async function checkGrammar() {
    const text = await getSelectedText();
    if (!text || text.trim().length === 0) {
        showResult('❌ 请先选中需要检查的文本', true);
        return;
    }
    
    const prompt = `请检查以下文本的语法、拼写和标点符号。列出所有问题和修正建议。

格式如下:
【问题 1】
原文: ...
错误: ...
修正: ...
原因: ...

如果没有问题，请说"✓ 文本完美！"`;
    
    const result = await callDeepSeekAPI(text, prompt);
    if (result) {
        showResult(`✓ 语法检查结果\n\n${result}`);
    }
}

// 改进内容
async function improveContent() {
    const text = await getSelectedText();
    if (!text || text.trim().length === 0) {
        showResult('❌ 请先选中需要改进的文本', true);
        return;
    }
    
    const prompt = `请改进以下文本的清晰度、简洁性和可读性。保持原意但使表达更好。

请提供:
1. 原文中的问题
2. 改进版本
3. 为什么这样改进更好

如果文本已经很好，请说"✓ 文本质量很高！"`;
    
    const result = await callDeepSeekAPI(text, prompt);
    if (result) {
        showResult(`↻ 改进建议\n\n${result}`);
    }
}

// 创意写作
async function generateText() {
    const text = await getSelectedText();
    if (!text || text.trim().length === 0) {
        showResult('❌ 请先选中主题或开头文本', true);
        return;
    }
    
    const prompt = `基于以下内容，生成创意、高质量的文本续写或扩展版本。保持专业的语气，确保内容连贯有趣。

提供 2-3 个不同的选项供选择:`;
    
    const result = await callDeepSeekAPI(text, prompt);
    if (result) {
        showResult(`✎ 创意写作建议\n\n${result}`);
    }
}

// 总结文本
async function summarizeText() {
    const text = await getSelectedText();
    if (!text || text.trim().length === 0) {
        showResult('❌ 请先选中需要总结的文本', true);
        return;
    }
    
    const prompt = `请为以下文本创建简洁清晰的总结。

要求:
- 保留主要要点
- 长度为原文的 1/3 左右
- 用项目符号格式
- 易于理解

总结:`;
    
    const result = await callDeepSeekAPI(text, prompt);
    if (result) {
        showResult(`📊 文本总结\n\n${result}`);
    }
}

// 数据分析
async function analyzeData() {
    const text = await getSelectedText();
    if (!text || text.trim().length === 0) {
        showResult('❌ 请先选中需要分析的数据或表格', true);
        return;
    }
    
    const prompt = `请分析以下数据或表格，提供:

1. 主要发现和趋势
2. 关键模式
3. 实际见解或建议
4. 任何异常值或需要注意的问题

使用清晰的格式和项目符号列表。`;
    
    const result = await callDeepSeekAPI(text, prompt);
    if (result) {
        showResult(`📈 数据分析结果\n\n${result}`);
    }
}

// 页面加载时恢复设置
document.addEventListener('DOMContentLoaded', function() {
    const savedApiKey = localStorage.getItem('deepseekApiKey');
    if (savedApiKey) {
        document.getElementById('apiKey').value = savedApiKey;
    }
});
