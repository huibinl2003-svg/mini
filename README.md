# Windows exe 云端打包模板（本地离线版）

这是一套现成的模板：你在自己电脑上不用装虚拟机，推送代码到 GitHub，微软的 Windows 服务器就自动帮你打出 exe，下载下来的 exe 是本地离线的，断网也能跑。

## 里面有什么

- `main.py`：示例程序，本地 tkinter 界面，不联网。换成你自己的源码就行，文件名保持 `main.py`。
- `requirements.txt`：打包要装的库，只写你真正用到的，比如 `openpyxl`、`python-docx`。
- `build.bat`：在你自己 Windows 电脑上双击就能本地打包，生成 `dist\MiniAssistant.exe`。
- `.github/workflows/build-windows-exe.yml`：GitHub Actions 配置，推送到 `main` 分支或手动点一下，就在 `windows-latest` 上用 Python 3.13 + PyInstaller 打包。

## 怎么用（云端打包）

1. 在 GitHub 新建一个仓库（私有也可以）。
2. 把这个模板里的全部文件传上去，把 `main.py` 换成你自己的程序源码，`requirements.txt` 按需改好。
3. 打开仓库的 Actions，找到 `Build Windows exe`，点 Run workflow，或者直接推送一次到 `main`。
4. 等它跑完，在那次运行的 Artifacts 里下载 `MiniAssistant-windows`，里面就是 `MiniAssistant.exe`，保存 30 天。

## 怎么用（本地打包）

在你自己的 Windows 电脑上，装好 Python 3.13，双击 `build.bat`，等它跑完，exe 在 `dist\MiniAssistant.exe`。

## 还缺什么

- 你的程序源码：模板里的 `main.py` 只是示例，Mini Assistant 的源码还在你自己电脑上，没有传给我。
- 一个 GitHub 账号和仓库：我这边没有你的 GitHub 账号，建仓库、推送代码这一步得你自己在浏览器里做，或者把源码发我，我帮你把文件整理好再给你。

## 注意

- 打出来的 exe 是离线的，不联网、不登录、不传数据，U盘拷到厂里没网的电脑上也能跑。
- 每次改完源码，重新推送一次，就会重新打包一个新的 exe，旧的不会自动覆盖。
