# 新教材词汇与学习资料

本仓库保存各年级教材词汇、教材 PDF、图片、音视频、处理脚本和生成的词汇手册。
大文件使用 Git LFS 保存。

## 下载和恢复

首次使用时，需要安装 Git 和 Git LFS。
macOS 可以通过 Homebrew 安装 Git LFS：

```bash
brew install git-lfs
```

进入希望保存项目的父目录，启用 Git LFS，然后克隆仓库：

```bash
git lfs install
git clone https://github.com/CaoYuhaoCarl/new_textbook_vocabulary.git
cd new_textbook_vocabulary
```

`git clone` 会自动创建项目目录和 Git 仓库，无需提前运行 `git init`。
启用 Git LFS 后，克隆时会自动下载大文件。
如果大文件没有下载完整，进入项目目录后执行：

```bash
git lfs pull
```

建议使用以上命令恢复完整资料。
GitHub 的 Download ZIP 可能只包含 LFS 指针文件，而不包含实际的大文件。

## 目录说明

- `7a`、`7b`、`8a`、`8b`、`9a`、`9b`：各年级教材和学习资料。
- 根目录的 CSV、Markdown、Excel 文件：词汇数据和整理结果。
- `scripts`：词汇分类和手册生成脚本。
- `outputs`：生成的词汇手册、分类结果和排版检查文件。

部分文件以 `.qkdownloading` 结尾，属于备份时尚未完成的下载文件。

## 删除本地目录前

修改资料后，需要提交并推送到 GitHub，才能在删除本地后恢复这些修改。
确认普通文件和 Git LFS 大文件都已上传，并验证可以从远程完整恢复后，再删除本地目录。
`.DS_Store` 是 macOS 目录元数据，不纳入备份。
