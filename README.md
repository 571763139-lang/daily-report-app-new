# 生产日报智能生成与校验系统 (无服务器 GitHub Pages 版)

本系统采用 **纯前端 WebAssembly Python** 技术构建。
通过将 Python 核心逻辑（`openpyxl`、`pandas` 和正则匹配）编译并在浏览器本地执行，本系统做到了 **100% 零服务器成本、100% 本地隐私安全**。您只需部署一个静态网页即可在全网各终端（包括手机端）无感使用。

---

## 💻 本地一键启动方式

由于现代浏览器的安全策略（CORS），直接双击打开 `index.html`（以 `file://` 协议运行）会因拦截 `fetch` 请求导致无法拉取资源。

1. **一键启动**：
   - 双击运行根目录下的 [一键本地启动.bat](file:///D:/ai项目/新日报生成系统/一键本地启动.bat)。
   - 脚本会调用 Python 启动一个微型静态 Web 服务器。
2. **访问系统**：
   - 打开浏览器，访问：`http://127.0.0.1:5000` 即可开始使用。

---

## ☁️ 公网全网部署流程 (GitHub Pages + Supabase 静态版)

### 第一步：在 Supabase 托管数据源模板 (可选，免上传)

如果您希望像您最初构想的那样，把“新模板数据源”托管在 Supabase 上，以便于脱离源码随时更改，请按以下步骤操作：

1. **注册与创建**：
   - 登录 [Supabase 官网](https://supabase.com/) 创建一个免费项目。
2. **创建 Storage 存储桶**：
   - 点击左侧菜单的 **Storage**。
   - 点击 **New Bucket**，将 Bucket Name 命名为 `templates`。
   - **重要**：打开 **Public Bucket** 选项（使其成为公共存储桶，方便免鉴权直接拉取），点击 Create。
3. **上传模板文件**：
   - 双击进入刚刚创建的 `templates` 存储桶，点击 **Upload**。
   - 将项目根目录下的 [新模板.xlsx](file:///D:/ai项目/新日报生成系统/新模板.xlsx) 上传上去。
4. **获取公共 URL**：
   - 上传成功后，点击文件右侧的 `...` 菜单，选择 **Copy URL**。
   - 您会获得一个形如 `https://your-id.supabase.co/storage/v1/object/public/templates/新模板.xlsx` 的公共下载链接。
5. **配置网页代码**：
   - 打开本地的 [static/main.js](file:///D:/ai项目/新日报生成系统/static/main.js)。
   - 将第 8 行的 `TEMPLATE_URL` 替换为您复制的 Supabase 公共链接：
     ```javascript
     const TEMPLATE_URL = "https://your-id.supabase.co/storage/v1/object/public/templates/新模板.xlsx";
     ```

> **💡 提示：**
> 如果您不想折腾 Supabase，可以直接在 `main.js` 里保持默认配置：
> `const TEMPLATE_URL = "./新模板.xlsx";`
> 这样，网页在初始化时会自动直接拉取您一同推送到 GitHub 仓库里的模板，完全省去了 Supabase 的注册配置！

---

### 第二步：一键发布到 GitHub Pages 静态网站 (永久免费)

1. **GitHub 准备**：
   - 登录您的 GitHub 账号，创建一个新的公开（Public）或私有（Private）的仓库。
   - 将本项目根目录下的所有代码及文件夹（包括 `index.html`、`static/` 文件夹以及 `新模板.xlsx`）全部提交并 push 到您的该 GitHub 仓库中。
2. **开启 GitHub Pages**：
   - 进入该仓库的 **Settings** -> **Pages**。
   - 在 **Build and deployment** 部分，Source 选择 **Deploy from a branch**。
   - Branch 选择 **main** (或 master) 分支，目录选择 **/ (root)**。
   - 点击 **Save**。
3. **部署上线**：
   - 稍等约 1 分钟，GitHub 会为您自动生成一个专属的公网链接（如 `https://your-username.github.io/repo-name/`）。
   - 打开该网址，网页便会在您的浏览器里利用 WebAssembly 技术自动启动 Python。环境初始化完成后即可在手机或电脑上完美使用！
