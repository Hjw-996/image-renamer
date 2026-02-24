# 图片重命名工具 - 云端打包指南

如果您无法在本地 Windows 电脑上进行打包，可以使用 GitHub Actions 在云端自动生成 `.exe` 文件。

## 步骤 1：上传代码到 GitHub
1.  登录您的 GitHub 账号（如果没有请注册一个）。
2.  创建一个新的仓库（Repository），例如命名为 `image-renamer`。
3.  将本文件夹中的所有文件上传到该仓库中。
    *   **关键文件**：`image_renamer.py` 和 `.github/workflows/build.yml`。
    *   **注意**：`.github` 文件夹必须保留其结构（即 `.github` 里面有 `workflows`，`workflows` 里面有 `build.yml`）。

## 步骤 2：等待自动打包
1.  上传完成后，点击仓库页面上方的 **Actions** 标签。
2.  您会看到一个名为 **Build Windows EXE** 的工作流正在运行（黄色旋转图标）。
3.  等待几分钟，直到图标变成绿色对勾，表示打包成功。

## 步骤 3：下载 EXE 文件
1.  点击那个绿色的 **Build Windows EXE** 任务。
2.  向下滚动到页面底部的 **Artifacts** 区域。
3.  点击 **Windows-EXE**。
4.  GitHub 会自动下载一个 `.zip` 压缩包，解压后即可得到可以在 Windows 上运行的 `图片重命名工具.exe`。

---
**提示**：
*   该服务是完全免费的。
*   生成的 EXE 文件可以在任何 Windows 10/11 电脑上运行，无需安装 Python。
