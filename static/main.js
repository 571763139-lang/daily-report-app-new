// ================= 配置后端 API 基础路径 =================
// 本地开发测试时，使用 "http://127.0.0.1:5000"
// 部署到 Koyeb / Render 等公网后，请在此处填写您的公网 API 网址（例如 "https://daily-report-api.koyeb.app"）
const API_BASE = "http://127.0.0.1:5000";
// ========================================================

document.addEventListener('DOMContentLoaded', () => {
    // 1. 获取 DOM 节点
    const form = document.getElementById('report-form');
    const misInput = document.getElementById('mis_file');
    const weighingInput = document.getElementById('weighing_file');
    const misZone = document.getElementById('mis-upload-zone');
    const weighingZone = document.getElementById('weighing-upload-zone');
    const misFileName = document.getElementById('mis-file-name');
    const weighingFileName = document.getElementById('weighing-file-name');
    
    const addCheckBtn = document.getElementById('add-check-btn');
    const checkListContainer = document.getElementById('check-list-container');
    
    const resultPlaceholder = document.getElementById('result-placeholder');
    const resultContent = document.getElementById('result-content');
    const verificationCard = document.getElementById('verification-card');
    const vIcon = document.getElementById('v-icon');
    const vStatusTitle = document.getElementById('v-status-title');
    const vStatusDesc = document.getElementById('v-status-desc');
    
    const zcPreviewList = document.getElementById('zc-preview-list');
    const byPreviewList = document.getElementById('by-preview-list');
    const sewagePreviewList = document.getElementById('sewage-preview-list');
    const hfComparisonGrid = document.getElementById('hf-comparison-grid');
    const downloadLink = document.getElementById('download-link');
    
    const loadingOverlay = document.getElementById('loading-overlay');
    const toast = document.getElementById('toast');

    // 2. 特殊字符序号列表 (① ~ ⑥)
    const numSymbols = ["①", "②", "③", "④", "⑤", "⑥"];
    let checkItemsCount = 0;

    // 3. 初始化默认添加一条检查项
    createCheckItem("无");

    // 4. 文件上传拖拽/点击交互
    setupFileZone(misZone, misInput, misFileName, "MIS数据");
    setupFileZone(weighingZone, weighingInput, weighingFileName, "称重数据");

    function setupFileZone(zone, input, nameElement, defaultLabel) {
        // 点击触发
        zone.addEventListener('click', () => input.click());

        // 选择文件后显示
        input.addEventListener('change', () => {
            if (input.files.length > 0) {
                const file = input.files[0];
                nameElement.textContent = file.name;
                zone.classList.add('active');
            } else {
                nameElement.textContent = `点击或拖拽上传 (.xlsx)`;
                zone.classList.remove('active');
            }
        });

        // 拖拽相关
        ['dragenter', 'dragover'].forEach(eventName => {
            zone.addEventListener(eventName, (e) => {
                e.preventDefault();
                zone.style.borderColor = 'var(--secondary)';
                zone.style.background = 'rgba(0, 240, 255, 0.04)';
            }, false);
        });

        ['dragleave', 'drop'].forEach(eventName => {
            zone.addEventListener(eventName, (e) => {
                e.preventDefault();
                zone.style.borderColor = '';
                zone.style.background = '';
            }, false);
        });

        zone.addEventListener('drop', (e) => {
            const dt = e.dataTransfer;
            const files = dt.files;
            if (files.length > 0) {
                input.files = files;
                nameElement.textContent = files[0].name;
                zone.classList.add('active');
            }
        });
    }

    // 5. 动态检查项列表管理
    addCheckBtn.addEventListener('click', () => {
        if (checkItemsCount >= 6) {
            showToast("检查内容最多只能添加 6 条！");
            return;
        }
        createCheckItem("");
    });

    function createCheckItem(initialValue = "") {
        checkItemsCount++;
        
        const checkItemDiv = document.createElement('div');
        checkItemDiv.className = 'check-item';
        
        const numLabel = document.createElement('span');
        numLabel.className = 'check-item-num';
        numLabel.textContent = numSymbols[checkItemsCount - 1];
        
        const input = document.createElement('input');
        input.type = 'text';
        input.name = 'check-content';
        input.placeholder = `请输入检查项 ${checkItemsCount} 内容...`;
        input.value = initialValue;
        
        const removeBtn = document.createElement('button');
        removeBtn.type = 'button';
        removeBtn.className = 'btn-remove-check';
        removeBtn.innerHTML = '<i class="fa-solid fa-trash-can"></i>';
        
        removeBtn.addEventListener('click', () => {
            checkItemDiv.remove();
            checkItemsCount--;
            rebuildCheckNums();
        });

        checkItemDiv.appendChild(numLabel);
        checkItemDiv.appendChild(input);
        checkItemDiv.appendChild(removeBtn);
        
        checkListContainer.appendChild(checkItemDiv);
    }

    function rebuildCheckNums() {
        const items = checkListContainer.querySelectorAll('.check-item');
        items.forEach((item, index) => {
            const numLabel = item.querySelector('.check-item-num');
            numLabel.textContent = numSymbols[index];
            const input = item.querySelector('input');
            input.placeholder = `请输入检查项 ${index + 1} 内容...`;
        });
    }

    // 6. 弹出式消息 Toast
    function showToast(message, duration = 3500) {
        toast.textContent = message;
        toast.classList.add('show');
        setTimeout(() => {
            toast.classList.remove('show');
        }, duration);
    }

    // 7. 表单提交与核心逻辑对接
    form.addEventListener('submit', async (e) => {
        e.preventDefault();

        // 基础验证
        if (!misInput.files.length) {
            showToast("⚠️ 请上传 MIS数据 文件！");
            return;
        }
        if (!weighingInput.files.length) {
            showToast("⚠️ 请上传 称重数据 文件！");
            return;
        }

        // 构建 FormData
        const formData = new FormData();
        formData.append('mis_file', misInput.files[0]);
        formData.append('weighing_file', weighingInput.files[0]);
        formData.append('baiyun_text', document.getElementById('baiyun_text').value);
        formData.append('zengcheng_text', document.getElementById('zengcheng_text').value);
        formData.append('huanfu_text', document.getElementById('huanfu_text').value);

        // 收集检查项数据
        const checkInputs = checkListContainer.querySelectorAll('input[name="check-content"]');
        const checksArray = [];
        checkInputs.forEach(input => {
            const val = input.value.trim();
            if (val && val !== '') {
                checksArray.push(val);
            }
        });
        formData.append('checks', JSON.stringify(checksArray));

        // 显示遮罩
        loadingOverlay.classList.add('active');

        try {
            // 对接跨域 API 地址
            const response = await fetch(`${API_BASE}/api/generate`, {
                method: 'POST',
                body: formData
            });

            const result = await response.json();
            loadingOverlay.classList.remove('active');

            if (result.success) {
                // 成功生成，渲染结果
                renderResult(result);
                showToast("🎉 日报表已成功生成并校验完成！");
            } else {
                showToast(`❌ 生成失败: ${result.error}`);
            }
        } catch (error) {
            loadingOverlay.classList.remove('active');
            showToast("❌ 网络连接或服务器响应异常！");
            console.error(error);
        }
    });

    // 8. 渲染分析和校验报告
    function renderResult(data) {
        // 隐藏占位符，显示数据区
        resultPlaceholder.style.display = 'none';
        resultContent.style.display = 'flex';

        // 1) 渲染校验卡片
        const cr = data.check_result;
        verificationCard.className = `verification-card ${cr.status}`;
        if (cr.status === 'success') {
            vIcon.className = 'fa-solid fa-circle-check';
            vStatusTitle.textContent = '校验通过';
            vStatusDesc.textContent = cr.message;
        } else {
            vIcon.className = 'fa-solid fa-circle-exclamation';
            vStatusTitle.textContent = '数据校验异常';
            vStatusDesc.textContent = cr.message;
        }

        // 2) 渲染增城建废数据
        const zc = data.parsed_data.zengcheng;
        zcPreviewList.innerHTML = `
            <li><span class="label">装修垃圾进厂量</span><span class="value">${zc['装修垃圾']} 吨</span></li>
            <li><span class="label">再生水稳量</span><span class="value">${zc['再生水稳']} 吨</span></li>
            <li><span class="label">再生石粉量</span><span class="value highlight">${zc['再生石粉']} 吨</span></li>
            <li><span class="label">砖渣出厂量</span><span class="value">${zc['砖渣']} 吨</span></li>
            <li><span class="label">混凝土块出厂量</span><span class="value">${zc['混凝土块']} 吨</span></li>
        `;

        // 3) 渲染白云建废数据
        const by = data.parsed_data.baiyun;
        byPreviewList.innerHTML = `
            <li><span class="label">进厂总量 (石渣/水泥头/原生石和)</span><span class="value">${by.total.toFixed(2)} 吨</span></li>
            <li><span class="label">原生混凝土</span><span class="value">${by.hunningtu} 立方米</span></li>
            <li><span class="label">再生1-2石</span><span class="value">${by.zaisheng12} 吨</span></li>
            <li><span class="label">再生水稳</span><span class="value">${by.zaishengshuiwen} 吨</span></li>
        `;

        // 4) 渲染称重数据
        const sew = data.parsed_data.sewage;
        sewagePreviewList.innerHTML = `
            <li><span class="label">污水外运总量</span><span class="value highlight">${sew.total.toFixed(2)} 吨</span></li>
            <li><span class="label">其中去环服</span><span class="value">${sew.to_hf.toFixed(2)} 吨</span></li>
            <li><span class="label">去电厂</span><span class="value">${sew.to_power.toFixed(2)} 吨</span></li>
        `;

        // 5) 渲染环服污水详情与各项统计
        const hf = data.parsed_data.huanfu;
        
        hfComparisonGrid.innerHTML = `
            <div class="comp-box highlight">
                <span class="title">文本总接收量</span>
                <span class="num">${hf['总量'].toFixed(2)} 吨</span>
            </div>
            <div class="comp-box success">
                <span class="title">称重算出去环服</span>
                <span class="num">${sew.to_hf.toFixed(2)} 吨</span>
            </div>
            <div class="comp-box">
                <span class="title">白云区接收</span>
                <span class="num">${hf['白云区'].toFixed(2)} 吨</span>
            </div>
            <div class="comp-box">
                <span class="title">人和镇接收</span>
                <span class="num">${hf['人和镇'].toFixed(2)} 吨</span>
            </div>
            <div class="comp-box">
                <span class="title">太和镇接收</span>
                <span class="num">${hf['太和镇'].toFixed(2)} 吨</span>
            </div>
            <div class="comp-box">
                <span class="title">云埔街道接收</span>
                <span class="num">${hf['云埔街道'].toFixed(2)} 吨</span>
            </div>
            <div class="comp-box">
                <span class="title">九龙镇接收</span>
                <span class="num">${hf['九龙镇'].toFixed(2)} 吨</span>
            </div>
            <div class="comp-box">
                <span class="title">天河城管接收</span>
                <span class="num">${hf['天河城管'].toFixed(2)} 吨</span>
            </div>
        `;

        // 6) 更新跨域下载地址
        downloadLink.href = `${API_BASE}/api/download/${data.file_id}`;
    }
});
