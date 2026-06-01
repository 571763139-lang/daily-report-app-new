// ================= 配置模板数据源地址 =================
// 1. (推荐) 如果您希望将模板嵌入在 Supabase：
//    请在 Supabase 创建公共 Bucket（如 templates），并上传 "新模板.xlsx"
//    将下方地址修改为您的 Supabase 该文件公共访问 URL：
//    const TEMPLATE_URL = "https://your-project-id.supabase.co/storage/v1/object/public/templates/新模板.xlsx";
// 2. 如果您希望将模板和网页源码一同打包在 GitHub 仓库中：
//    请将下方地址修改为相对路径 './新模板.xlsx' 即可：
const TEMPLATE_URL = "./新模板.xlsx";
// ======================================================

document.addEventListener('DOMContentLoaded', async () => {
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

    // 进度条及初始化指示器 DOM
    const envProgress = document.getElementById('env-progress');
    const envText = document.getElementById('env-text');
    const envIcon = document.getElementById('env-icon');

    // 特殊字符序号列表 (① ~ ⑥)
    const numSymbols = ["①", "②", "③", "④", "⑤", "⑥"];
    let checkItemsCount = 0;

    // 核心的 WebAssembly Python 逻辑代码及 Pyodide 实例
    let pyodideInstance = null;
    let pythonCoreCode = "";

    // 网页启动时开始异步加载环境
    initWasmEnvironment();

    // 2. 初始化 Pyodide 环境及抓取数据源
    async function initWasmEnvironment() {
        try {
            // 步骤 A: 拉取模板数据与逻辑代码
            updateProgress(15, "正在连接数据源拉取新模板.xlsx...");
            const templateResp = await fetch(TEMPLATE_URL);
            if (!templateResp.ok) throw new Error("无法从指定的 URL 拉取新模板，请核对 TEMPLATE_URL 配置！");
            const templateBuffer = await templateResp.arrayBuffer();

            const logicResp = await fetch('static/core_logic.py');
            if (!logicResp.ok) throw new Error("无法读取本地核心 Python 处理逻辑脚本！");
            pythonCoreCode = await logicResp.text();

            // 步骤 B: 启动 Pyodide Wasm (使用 Fastly 官方加速 CDN，国内秒加载且无重定向限制)
            updateProgress(35, "正在启动浏览器 Python 虚拟环境 (1/3)...");
            pyodideInstance = await loadPyodide({
                indexURL: "https://fastly.jsdelivr.net/pyodide/v0.23.4/full/"
            });

            // 步骤 C: 安装依赖包
            updateProgress(65, "正在装载 Pandas 和 OpenPyXL 依赖库 (2/3)...");
            await pyodideInstance.loadPackage(["pandas", "openpyxl"]);

            // 步骤 D: 挂载模板文件
            updateProgress(90, "正在同步文件和模板 (3/3)...");
            pyodideInstance.FS.writeFile("新模板.xlsx", new Uint8Array(templateBuffer));

            // 初始化成功，解锁表单
            updateProgress(100, "环境装载成功！系统已就绪，正在本地极速运行中。");
            envIcon.className = "fa-solid fa-circle-check text-success";
            unlockForm();

            // 初始化添加第一条检查项
            createCheckItem("无");

        } catch (error) {
            console.error(error);
            envIcon.className = "fa-solid fa-circle-xmark text-danger";
            envText.innerHTML = `<span class="text-danger">环境装载失败: ${error.message}</span>`;
            showToast("❌ 初始化环境失败，请核对您的网络或数据源链接！");
        }
    }

    function updateProgress(percentage, text) {
        envProgress.style.width = `${percentage}%`;
        envText.textContent = text;
    }

    function unlockForm() {
        // 解锁所有被 disabled 的输入控件
        misInput.disabled = false;
        weighingInput.disabled = false;
        document.getElementById('zengcheng_text').disabled = false;
        document.getElementById('baiyun_text').disabled = false;
        document.getElementById('huanfu_text').disabled = false;
        addCheckBtn.disabled = false;
        document.getElementById('submit-btn').disabled = false;

        misFileName.textContent = "点击或拖拽上传 (.xlsx)";
        weighingFileName.textContent = "点击或拖拽上传 (.xlsx)";
    }

    // 3. 文件上传拖拽/点击交互
    setupFileZone(misZone, misInput, misFileName);
    setupFileZone(weighingZone, weighingInput, weighingFileName);

    function setupFileZone(zone, input, nameElement) {
        zone.addEventListener('click', () => {
            if (!input.disabled) input.click();
        });

        input.addEventListener('change', () => {
            if (input.files.length > 0) {
                nameElement.textContent = input.files[0].name;
                zone.classList.add('active');
            } else {
                nameElement.textContent = "点击或拖拽上传 (.xlsx)";
                zone.classList.remove('active');
            }
        });

        ['dragenter', 'dragover'].forEach(eventName => {
            zone.addEventListener(eventName, (e) => {
                e.preventDefault();
                if (input.disabled) return;
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
            if (input.disabled) return;
            const dt = e.dataTransfer;
            const files = dt.files;
            if (files.length > 0) {
                input.files = files;
                nameElement.textContent = files[0].name;
                zone.classList.add('active');
            }
        });
    }

    // 4. 动态检查项列表管理
    addCheckBtn.addEventListener('click', () => {
        if (addCheckBtn.disabled) return;
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

    // 5. Toast 消息提示
    function showToast(message, duration = 3500) {
        toast.textContent = message;
        toast.classList.add('show');
        setTimeout(() => {
            toast.classList.remove('show');
        }, duration);
    }

    // 6. 表单提交：通过 WebAssembly 执行 Python 逻辑并直接输出结果
    form.addEventListener('submit', async (e) => {
        e.preventDefault();

        if (!pyodideInstance) {
            showToast("⚠️ Python 运行环境尚未准备就绪，请稍候！");
            return;
        }

        // 文件非空验证
        if (!misInput.files.length || !weighingInput.files.length) {
            showToast("⚠️ 请确保上传了 MIS数据 和 称重数据！");
            return;
        }

        // 展现加载动画
        loadingOverlay.classList.add('active');

        try {
            // 读取用户上传的 Excel 文件为 ArrayBuffer
            const misBuffer = await misInput.files[0].arrayBuffer();
            const weighingBuffer = await weighingInput.files[0].arrayBuffer();

            // 写入 Pyodide 虚拟文件系统
            pyodideInstance.FS.writeFile("mis_file.xlsx", new Uint8Array(misBuffer));
            pyodideInstance.FS.writeFile("weighing_file.xlsx", new Uint8Array(weighingBuffer));

            // 将文本输入变量注入 Python 全局作用域
            pyodideInstance.globals.set("baiyun_text", document.getElementById('baiyun_text').value);
            pyodideInstance.globals.set("zengcheng_text", document.getElementById('zengcheng_text').value);
            pyodideInstance.globals.set("huanfu_text", document.getElementById('huanfu_text').value);

            // 收集检查项数据
            const checkInputs = checkListContainer.querySelectorAll('input[name="check-content"]');
            const checksArray = [];
            checkInputs.forEach(input => {
                const val = input.value.trim();
                if (val && val !== '') checksArray.push(val);
            });
            pyodideInstance.globals.set("checks_json", JSON.stringify(checksArray));

            // 异步执行 Python 脚本
            await pyodideInstance.runPythonAsync(pythonCoreCode);

            // 从 Python 作用域获取返回的 JSON 字符串结果
            const outputJsonStr = pyodideInstance.globals.get("output_json_str");
            const result = JSON.parse(outputJsonStr);

            if (result.success) {
                // 从虚拟文件系统中读取生成好的 excel 二进制文件
                const outputExcelBytes = pyodideInstance.FS.readFile("output_file.xlsx");
                
                // 将二进制转换为本地下载 Blob 流
                const blob = new Blob([outputExcelBytes], { 
                    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" 
                });
                const blobUrl = URL.createObjectURL(blob);

                // 更新前端界面结果面板
                renderResult(result, blobUrl);
                
                loadingOverlay.classList.remove('active');
                showToast("🎉 日报表已由本地 Python 成功生成并校验完成！");
            } else {
                throw new Error("Python 内部生成失败。");
            }

        } catch (error) {
            loadingOverlay.classList.remove('active');
            showToast(`❌ 数据解析或写入错误: ${error.message}`);
            console.error(error);
        }
    });

    // 7. 渲染结果仪表盘
    function renderResult(data, blobUrl) {
        // 隐藏占位符，显示数据区
        resultPlaceholder.style.display = 'none';
        resultContent.style.display = 'flex';

        // 1) 校验预警卡片
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

        // 2) 增城建废提取结果
        const zc = data.parsed_data.zengcheng;
        zcPreviewList.innerHTML = `
            <li><span class="label">装修垃圾进厂量</span><span class="value">${zc['装修垃圾']} 吨</span></li>
            <li><span class="label">再生水稳量</span><span class="value">${zc['再生水稳']} 吨</span></li>
            <li><span class="label">再生石粉量</span><span class="value highlight">${zc['再生石粉']} 吨</span></li>
            <li><span class="label">砖渣出厂量</span><span class="value">${zc['砖渣']} 吨</span></li>
            <li><span class="label">混凝土块出厂量</span><span class="value">${zc['混凝土块']} 吨</span></li>
        `;

        // 3) 白云建废提取结果
        const by = data.parsed_data.baiyun;
        byPreviewList.innerHTML = `
            <li><span class="label">进厂总量 (石渣/水泥头/原生石和)</span><span class="value">${by.total.toFixed(2)} 吨</span></li>
            <li><span class="label">原生混凝土</span><span class="value">${by.hunningtu} 立方米</span></li>
            <li><span class="label">再生1-2石</span><span class="value">${by.zaisheng12} 吨</span></li>
            <li><span class="label">再生水稳</span><span class="value">${by.zaishengshuiwen} 吨</span></li>
        `;

        // 4) 称重污水外运
        const sew = data.parsed_data.sewage;
        sewagePreviewList.innerHTML = `
            <li><span class="label">污水外运总量</span><span class="value highlight">${sew.total.toFixed(2)} 吨</span></li>
            <li><span class="label">其中去环服</span><span class="value">${sew.to_hf.toFixed(2)} 吨</span></li>
            <li><span class="label">去电厂</span><span class="value">${sew.to_power.toFixed(2)} 吨</span></li>
        `;

        // 5) 环服污水详情对比
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

        // 6) 更新本地下载二进制
        downloadLink.href = blobUrl;
        downloadLink.download = data.filename;
    }
});
