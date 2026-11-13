let pyodide = null;
let coreCode = '';
const VERSION = '5.0.0';
function progress(text, percent) { self.postMessage({type:'progress', text, percent}); }
async function resource(url, kind) {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), 30000);
    try {
        const response = await fetch(url, {signal:controller.signal});
        if (!response.ok) throw new Error(`无法读取${kind}（${response.status}），请检查发布文件是否齐全`);
        return kind === '模板' ? await response.arrayBuffer() : await response.text();
    } finally { clearTimeout(timer); }
}
self.onmessage = async ({data}) => {
    const {id, type} = data;
    try {
        if (type === 'init') {
            progress('正在加载模板和制表规则程序…', 15);
            const [template, code] = await Promise.all([resource(`../新模板.xlsx?v=${VERSION}`, '模板'),resource(`core_logic.py?v=${VERSION}`, '处理程序')]);
            coreCode = code;
            progress('正在后台启动计算环境，首次使用需下载依赖…', 35);
            importScripts('https://fastly.jsdelivr.net/pyodide/v0.23.4/full/pyodide.js');
            pyodide = await loadPyodide({indexURL:'https://fastly.jsdelivr.net/pyodide/v0.23.4/full/'});
            await pyodide.loadPackage('micropip');
            progress('正在加载 Excel 处理组件…', 70);
            const micropip = pyodide.pyimport('micropip');
            try { await micropip.install('openpyxl==3.1.5'); } finally { micropip.destroy(); }
            pyodide.FS.writeFile('新模板.xlsx', new Uint8Array(template));
            progress('计算环境已就绪，可导入数据生成日报。', 100);
            self.postMessage({id, ok:true});
        } else if (type === 'generate') {
            if (!pyodide) throw new Error('计算环境尚未就绪，请重新初始化');
            progress('正在读取 Excel、计算汇总并填入模板…', 100);
            try {
                pyodide.FS.writeFile('mis_file.xlsx', new Uint8Array(data.mis));
                pyodide.FS.writeFile('weighing_file.xlsx', new Uint8Array(data.weighing));
                for (const [name,value] of Object.entries(data.inputs)) pyodide.globals.set(name,value);
                pyodide.globals.set('output_json_str', '');
                await pyodide.runPythonAsync(coreCode);
                const result = JSON.parse(pyodide.globals.get('output_json_str'));
                const bytes = pyodide.FS.readFile('output_file.xlsx');
                self.postMessage({id,ok:true,result,excel:bytes.buffer},[bytes.buffer]);
            } finally {
                for (const name of ['mis_file.xlsx','weighing_file.xlsx','output_file.xlsx']) {
                    try { pyodide.FS.unlink(name); } catch (_) { /* 文件可能尚未创建 */ }
                }
                await pyodide.runPythonAsync('import gc; gc.collect()');
            }
        } else throw new Error('未知操作');
    } catch (error) { self.postMessage({id,ok:false,error:String(error.message || error)}); }
};
