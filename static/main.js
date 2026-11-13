'use strict';
document.addEventListener('DOMContentLoaded', () => {
    const $ = id => document.getElementById(id);
    const form = $('report-form'), submit = $('submit-btn');
    let worker = null, ready = false, busy = false, requestId = 0, blobUrl = null;
    let currentRun = null, toastTimer = null;
    const pending = new Map();
    function toast(message) {
        clearTimeout(toastTimer); $('toast').textContent = message; $('toast').classList.add('show');
        toastTimer = setTimeout(() => $('toast').classList.remove('show'), 4000);
    }
    function errorText(error) {
        const lines = String(error.message || error).trim().split('\n').filter(Boolean);
        return lines[lines.length - 1].replace(/^(?:ValueError|Error):\s*/, '');
    }
    function showError(error) { $('error-message').textContent = errorText(error); $('error-panel').hidden = false; }
    function controls() {
        for (const element of form.querySelectorAll('input,textarea,button')) element.disabled = !ready || busy;
        submit.setAttribute('aria-busy', String(busy)); $('task-status').hidden = !busy;
    }
    function clearResult() {
        if (blobUrl) URL.revokeObjectURL(blobUrl);
        blobUrl = null; $('download-link').removeAttribute('href');
        $('result-content').style.display = 'none'; $('result-placeholder').style.display = 'flex';
    }
    function stopWorker(reason = new Error('计算已取消')) {
        if (worker) worker.terminate(); worker = null; ready = false;
        for (const {reject,timer} of pending.values()) { clearTimeout(timer); reject(reason); }
        pending.clear(); controls();
    }
    function request(type, payload = {}, transfer = []) {
        return new Promise((resolve,reject) => {
            if (!worker) { reject(new Error('请先重新初始化计算环境')); return; }
            const id = ++requestId;
            const timer = setTimeout(() => {
                const error = new Error(type === 'init' ? '环境加载超时，请检查网络后点击重新初始化；手工制表规则仍可打开。' : '生成超过120秒，请核对文件大小或重新导出Excel后重试。');
                stopWorker(error); $('retry-env-btn').hidden = false;
            }, type === 'init' ? 180000 : 120000);
            pending.set(id,{resolve,reject,timer});
            try { worker.postMessage({id,type,...payload},transfer); }
            catch (error) { clearTimeout(timer); pending.delete(id); reject(error); }
        });
    }
    async function init() {
        $('retry-env-btn').hidden = true; $('error-panel').hidden = true; stopWorker();
        $('env-icon').className = 'fa-solid fa-circle-notch fa-spin text-secondary';
        $('env-text').textContent = '正在后台加载计算环境，期间可阅读手工制表规则…'; $('env-progress').style.width = '5%';
        try {
            if (location.protocol === 'file:') throw new Error('请使用“一键本地启动.bat”打开此版本，不要直接双击index.html。');
            worker = new Worker('static/report-worker.js?v=5.0.1');
            worker.onmessage = ({data}) => {
                if (data.type === 'progress') {
                    $('env-text').textContent = data.text; $('env-progress').style.width = `${data.percent}%`;
                    if (busy) $('task-message').textContent = data.text;
                    return;
                }
                const task = pending.get(data.id); if (!task) return;
                clearTimeout(task.timer); pending.delete(data.id);
                if (data.ok) task.resolve(data); else task.reject(new Error(data.error));
            };
            worker.onerror = event => {
                event.preventDefault();
                const error = new Error('后台计算环境异常，请点击重新初始化；页面仍可滚动，已填内容保留。');
                stopWorker(error); showError(error); $('retry-env-btn').hidden = false;
            };
            await request('init'); ready = true; $('env-icon').className = 'fa-solid fa-circle-check text-success';
            for (const name of ['mis','weighing']) if (!$(`${name}_file`).files.length) $(`${name}-file-name`).textContent = '选择本机Excel（.xlsx）';
        } catch (error) {
            stopWorker(error); $('env-icon').className = 'fa-solid fa-circle-xmark text-danger';
            $('env-text').textContent = errorText(error); $('retry-env-btn').hidden = false; showError(error);
        } finally { controls(); }
    }
    $('retry-env-btn').addEventListener('click',init);
    // 原生文件输入覆盖卡片，不从父节点递归触发 input.click()。
    for (const name of ['mis','weighing']) {
        const input = $(`${name}_file`), zone = $(`${name}-upload-zone`), label = $(`${name}-file-name`);
        input.addEventListener('change', () => {
            label.textContent = input.files[0]?.name || '选择本机Excel（.xlsx）';
            zone.classList.toggle('active',!!input.files.length); clearResult();
        });
        zone.addEventListener('dragover',event => event.preventDefault());
        zone.addEventListener('drop',event => {
            event.preventDefault(); if (input.disabled) return;
            if (event.dataTransfer.files.length !== 1) { showError(new Error('每个数据源只接受一个.xlsx文件')); return; }
            try { input.files = event.dataTransfer.files; input.dispatchEvent(new Event('change',{bubbles:true})); }
            catch (_) { showError(new Error('浏览器不支持拖入文件，请点击选择文件')); }
        });
    }
    function renumber() {
        $('check-list-container').querySelectorAll('.check-item').forEach((item,index) => {
            item.querySelector('.check-item-num').textContent = index + 1;
            const input = item.querySelector('input'); input.placeholder = `请输入检查项 ${index + 1} 内容…`;
            input.setAttribute('aria-label',`检查项 ${index + 1}`); item.querySelector('button').setAttribute('aria-label',`删除检查项 ${index + 1}`);
        });
    }
    function addCheck(value = '') {
        if ($('check-list-container').children.length >= 6) { toast('检查内容最多6条'); return; }
        const item = document.createElement('div'); item.className = 'check-item';
        const number = document.createElement('span'); number.className = 'check-item-num';
        const input = document.createElement('input'); input.type = 'text'; input.name = 'check-content'; input.value = value;
        const remove = document.createElement('button'); remove.type = 'button'; remove.className = 'btn-remove-check'; remove.textContent = '删除';
        remove.addEventListener('click',() => { item.remove(); renumber(); clearResult(); });
        item.append(number,input,remove); $('check-list-container').append(item); renumber();
    }
    addCheck('无'); controls();
    $('add-check-btn').addEventListener('click',() => { addCheck(); controls(); }); form.addEventListener('input',clearResult);
    function readFile(file, signal) {
        return new Promise((resolve,reject) => {
            const reader = new FileReader(); let timer;
            const cleanup = () => { clearTimeout(timer); signal.removeEventListener('abort',abort); };
            const abort = () => { reader.abort(); cleanup(); reject(new DOMException('已取消','AbortError')); };
            reader.onload = () => { cleanup(); resolve(reader.result); };
            reader.onerror = () => { cleanup(); reject(new Error(`无法读取“${file.name}”，请先将iCloud/网盘文件下载到本机再选择。`)); };
            timer = setTimeout(() => { cleanup(); reader.onload=null; reader.onerror=null; reader.abort(); reject(new Error(`读取“${file.name}”超时，请先下载到本机。`)); },30000);
            signal.addEventListener('abort',abort,{once:true}); if (signal.aborted) { abort(); return; }
            reader.readAsArrayBuffer(file);
        });
    }
    function validateFile(file, label) {
        if (!file || !/\.xlsx$/i.test(file.name)) throw new Error(`${label}请选择.xlsx文件，不支持.xls或直接改后缀的文件`);
        if (!file.size) throw new Error(`${label}是空文件，请重新下载`);
        if (file.size > 20*1024*1024) throw new Error(`${label}超过20MB，请精简无用行列后重新导出`);
    }
    function render(data, excel) {
        blobUrl = URL.createObjectURL(new Blob([excel],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));
        $('download-link').href = blobUrl; $('download-link').download = data.filename;
        $('result-placeholder').style.display = 'none'; $('result-content').style.display = 'flex';
        const warnings = data.warnings || [];
        $('verification-card').className = `verification-card ${warnings.length ? 'warning' : 'success'}`;
        $('v-icon').className = `fa-solid ${warnings.length ? 'fa-triangle-exclamation' : 'fa-circle-check'}`;
        $('v-status-title').textContent = warnings.length ? '已生成，有待核对项' : '计算完成';
        $('v-status-desc').textContent = warnings.length ? '请逐条核对下方提示，再使用日报。' : '污水数据已计算并写入表格，请与来源核对。';
        $('result-warnings').replaceChildren(...warnings.map(message => { const li=document.createElement('li'); li.textContent=message; return li; }));
        $('result-warnings').hidden = !warnings.length;
        const fill = (id,rows) => $(id).replaceChildren(...rows.map(([label,value,unit]) => {
            const li=document.createElement('li'), l=document.createElement('span'), v=document.createElement('span');
            l.className='label'; l.textContent=label; v.className='value'; v.textContent=`${value} ${unit || '吨'}`;
            li.append(l,v); return li;
        }));
        const p=data.parsed_data, z=p.zengcheng, b=p.baiyun, s=p.sewage;
        const source=data.weighing_source;
        $('weighing-source').textContent=source ? `已读取：${source.sheet} · 第${source.header_row}行表头 · ${source.records}条明细` : '';
        fill('fly-ash-preview-list',[['飞灰外运资源化',p.fly_ash]]);
        fill('zc-preview-list',Object.entries(z).map(([key,value])=>[key,value]));
        fill('by-preview-list',[['进厂总量（石渣+水泥头+原生石）',b.total.toFixed(2)],['原生混凝土',b.hunningtu,'立方米'],['再生1-2石',b.zaisheng12],['再生水稳',b.zaishengshuiwen]]);
        fill('sewage-preview-list',[['生活垃圾污水运输总量',s.total.toFixed(2)],['环境集团运输',s.by_huantou.toFixed(2)],['社会单位运输',s.by_social.toFixed(2)],['环服公司处置',s.to_huanfu.toFixed(2)],['各资源热力电厂处置',s.to_power.toFixed(2)]]);
    }
    form.addEventListener('submit',async event => {
        event.preventDefault(); if (busy || !ready || !form.reportValidity()) return;
        clearResult(); $('error-panel').hidden = true;
        const controller = new AbortController(); currentRun = controller;
        try {
            const mis=$('mis_file').files[0], weighing=$('weighing_file').files[0];
            validateFile(mis,'MIS数据'); validateFile(weighing,'称重数据');
            const flyAsh = Number($('fly_ash_amount').value);
            if ($('fly_ash_amount').value.trim()==='' || !Number.isFinite(flyAsh) || flyAsh<0) throw new Error('飞灰量请输入大于等于0的有效数字');
            const inputs = {fly_ash_amount:flyAsh,baiyun_text:$('baiyun_text').value,zengcheng_text:$('zengcheng_text').value,
                checks_json:JSON.stringify([...form.querySelectorAll('[name="check-content"]')].map(input=>input.value.trim()).filter(Boolean))};
            busy=true; controls(); $('task-message').textContent='正在读取本机文件，可继续滚动页面或查看规则…';
            const [misBytes,weighingBytes] = await Promise.all([readFile(mis,controller.signal),readFile(weighing,controller.signal)]);
            if (controller.signal.aborted) throw new DOMException('已取消','AbortError');
            const data=await request('generate',{mis:misBytes,weighing:weighingBytes,inputs},[misBytes,weighingBytes]);
            render(data.result,data.excel); toast('日报已生成，请核对预览后下载。');
        } catch (error) { controller.abort(); if (error.name !== 'AbortError') showError(error); }
        finally {
            busy=false; currentRun=null; controls();
            if (ready) $('env-text').textContent='计算环境已就绪，可修改数据后重新生成。';
        }
    });
    $('cancel-btn').addEventListener('click',() => {
        currentRun?.abort(); stopWorker(new DOMException('已取消','AbortError'));
        $('env-text').textContent='生成已取消，已填内容保留。点击重新初始化后可重试。'; $('retry-env-btn').hidden=false;
    });
    // 不改写 body 的 overflow，关闭说明后不会遗留滚动锁。
    const dialog=$('rules-dialog'); let savedFocus=null, savedScroll=0;
    $('rules-btn').addEventListener('click',() => {
        savedFocus=document.activeElement; savedScroll=window.scrollY;
        if (dialog.showModal) dialog.showModal(); else dialog.setAttribute('open','');
        $('rules-close').focus({preventScroll:true});
    });
    function restoreFocus() { savedFocus?.focus({preventScroll:true}); window.scrollTo(0,savedScroll); }
    const closeRules = () => { if (dialog.close) dialog.close(); else { dialog.removeAttribute('open'); restoreFocus(); } };
    dialog.addEventListener('close',restoreFocus);
    for (const button of dialog.querySelectorAll('[data-close-rules]')) button.addEventListener('click',closeRules);
    dialog.addEventListener('click',event => { if (event.target===dialog) closeRules(); });
    dialog.addEventListener('keydown',event => {
        if (event.key==='Escape') { event.preventDefault(); closeRules(); }
        if (event.key==='Tab') {
            const targets=[...dialog.querySelectorAll('button,a[href], [tabindex="0"]')], first=targets[0], last=targets[targets.length-1];
            if (event.shiftKey && document.activeElement===first) { event.preventDefault(); last.focus(); }
            else if (!event.shiftKey && document.activeElement===last) { event.preventDefault(); first.focus(); }
        }
    });
    window.addEventListener('pagehide',() => { stopWorker(); if (blobUrl) URL.revokeObjectURL(blobUrl); });
    window.addEventListener('pageshow',event => { if (event.persisted) { busy=false; clearResult(); init(); } });
    init();
});
