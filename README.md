# Hybrid Image to PPTX

将幻灯片截图、CAE 仿真结果页和技术架构图还原为可编辑的 PowerPoint。输入原图，输出 `.pptx`，保留原图的文字、数值、布局和技术关系。

- 标题、正文、节点框、连接线和公式使用 PowerPoint 原生对象。公式可在 PowerPoint 中编辑。
- CFD 云图、照片、Logo 和复杂三维图保留为独立图片，便于移动和替换。
- 连接线保留原图走向。移动节点后，需手动调整对应端点。

## 在 Codex 中安装与使用

在 Codex 中发送这条安装指令。

```text
用 $skill-installer 从 MingostemonHuang/Hybrid-image2pptx 的 main 分支安装技能。
仓库内路径是 .（根目录），安装名称为 hybrid-image2pptx。
```

macOS 和 Linux 用户也可以在终端手动安装。

```bash
git clone https://github.com/MingostemonHuang/Hybrid-image2pptx.git ~/.agents/skills/hybrid-image2pptx
```

上传原图后，用这条指令调用技能。

```text
用 $hybrid-image2pptx 把这张截图还原为可编辑 PPTX。
保留原图术语、数值和箭头关系，将 CFD 云图保留为独立图片。
```

技能没有出现在选择器时，重启 Codex。 [Codex 技能文档](https://learn.chatgpt.com/docs/build-skills)

## 在 ChatGPT 工作区中上传

符合条件的 ChatGPT Business、Enterprise、Healthcare 和 Edu 工作区可从本机上传技能，具体取决于功能开放情况和管理员设置。普通个人 ChatGPT 账号和自定义 GPT 没有从 GitHub 链接一键安装此技能的入口。 [ChatGPT Skills 帮助](https://help.openai.com/en/articles/20001066-skills-in-chatgpt)

1. 下载并解压仓库，把含有 `SKILL.md` 的目录命名为 `hybrid-image2pptx`。保留其中的 `agents/`、`references/`、`scripts/` 和 `requirements.txt`。 [下载源码 ZIP](https://github.com/MingostemonHuang/Hybrid-image2pptx/archive/refs/heads/main.zip)
2. 在 ChatGPT 侧栏打开 **Plugins → Skills → Create → Upload from your computer**，上传完整技能文件。如果界面要求 ZIP，将技能目录打包为 `hybrid-image2pptx.zip` 再上传。
3. 上传通过后，在对话中输入 `@` 选择 **Hybrid Image to PPTX**，附上原图并要求导出 `.pptx`。

导出需要 Python 和 `requirements.txt` 中的依赖。本地导出已验证，ChatGPT 工作区内的上传和导出需在目标环境中验证。

## 用 Python 构建 PPTX

在仓库根目录，按场景格式准备 `scene.json` 和原图。场景文件记录各元素的位置、内容以及图片的裁剪范围。 [场景格式](references/scene-format.md)

```bash
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python scripts/build_pptx.py scene.json output.pptx
.venv/bin/python scripts/inspect_pptx.py output.pptx
```

Windows 下用 `.venv\Scripts\python.exe` 执行虚拟环境中的命令。重建与验收规则见 [SKILL.md](SKILL.md)。
