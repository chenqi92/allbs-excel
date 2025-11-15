# GitHub Actions 自动发布配置

本目录包含了自动发布到 Maven Central 和 GitHub Packages 的 GitHub Actions 配置。

---

## 📁 文件说明

```
.github/
├── workflows/
│   ├── maven-publish.yml                    # 完整版 Workflow（包含详细步骤）
│   └── publish-on-version-change.yml        # 简化版 Workflow（推荐使用）
├── scripts/
│   └── setup-gpg.sh                         # GPG 密钥生成脚本
├── SETUP_GUIDE.md                           # 详细配置指南
└── README.md                                # 本文件
```

---

## 🚀 快速开始

### 1. 生成 GPG 密钥

运行自动化脚本：

```bash
chmod +x .github/scripts/setup-gpg.sh
./.github/scripts/setup-gpg.sh
```

或手动生成：

```bash
# 生成密钥
gpg --full-generate-key

# 导出私钥
gpg --armor --export-secret-keys YOUR_KEY_ID | base64 > gpg-private-key.txt
```

### 2. 配置 GitHub Secrets

在 GitHub 仓库中配置以下 4 个 Secrets：

| Secret 名称 | 说明 | 获取方式 |
|------------|------|---------|
| `MAVEN_USERNAME` | Maven Central 用户名 | [Sonatype Central](https://central.sonatype.com/) |
| `MAVEN_PASSWORD` | Maven Central 密码/Token | Sonatype Central → Generate Token |
| `GPG_PRIVATE_KEY` | GPG 私钥 (Base64) | 运行 setup-gpg.sh 脚本 |
| `GPG_PASSPHRASE` | GPG 密钥密码 | 生成 GPG 密钥时设置的密码 |

**配置步骤**:
1. 打开仓库 → `Settings` → `Secrets and variables` → `Actions`
2. 点击 `New repository secret`
3. 添加上述 4 个 Secrets

### 3. 修改版本号并发布

```bash
# 1. 修改 pom.xml 中的版本号
# 例如: 从 3.0.0 改为 3.0.1

# 2. 提交并推送
git add pom.xml
git commit -m "chore: bump version to 3.0.1"
git push origin main

# 3. GitHub Actions 会自动检测版本变化并发布
```

---

## 📖 详细文档

完整的配置指南请查看: [SETUP_GUIDE.md](SETUP_GUIDE.md)

---

## 🔄 Workflow 说明

### publish-on-version-change.yml (推荐)

**触发条件**:
- 推送到 `main` 分支
- 修改了 `pom.xml` 文件
- 版本号发生变化

**执行流程**:
1. ✅ 检测版本号是否变化
2. ✅ 编译和测试
3. ✅ 发布到 Maven Central
4. ✅ 发布到 GitHub Packages
5. ✅ 创建 GitHub Release (仅 Release 版本)
6. ✅ 创建 Git Tag (仅 Release 版本)

**特点**:
- 简洁高效
- 自动检测版本变化
- 支持 SNAPSHOT 和 Release 版本
- 自动创建 Release 和 Tag

### maven-publish.yml (完整版)

包含更详细的步骤和日志输出，适合调试和学习。

---

## 📦 发布流程

### Release 版本发布

```xml
<!-- pom.xml -->
<version>3.0.1</version>
```

```bash
git add pom.xml
git commit -m "chore: release version 3.0.1"
git push origin main
```

**结果**:
- ✅ 发布到 Maven Central
- ✅ 发布到 GitHub Packages
- ✅ 创建 GitHub Release
- ✅ 创建 Git Tag `v3.0.1`

### SNAPSHOT 版本发布

```xml
<!-- pom.xml -->
<version>3.0.2-SNAPSHOT</version>
```

```bash
git add pom.xml
git commit -m "chore: prepare for next development iteration"
git push origin main
```

**结果**:
- ✅ 发布到 Maven Central (SNAPSHOT 仓库)
- ✅ 发布到 GitHub Packages
- ❌ 不创建 GitHub Release
- ❌ 不创建 Git Tag

---

## 🔍 查看发布结果

### Maven Central
https://central.sonatype.com/artifact/cn.allbs/allbs-excel

### GitHub Packages
https://github.com/chenqi92/allbs-excel/packages

### GitHub Releases
https://github.com/chenqi92/allbs-excel/releases

---

## 🛠️ 故障排查

### 问题 1: Workflow 没有触发

**检查**:
- ✅ 是否推送到 `main` 分支
- ✅ 是否修改了 `pom.xml` 文件
- ✅ 版本号是否真的发生了变化

### 问题 2: GPG 签名失败

**检查**:
- ✅ `GPG_PRIVATE_KEY` 是否正确 Base64 编码
- ✅ `GPG_PASSPHRASE` 是否正确
- ✅ 公钥是否已上传到密钥服务器

### 问题 3: Maven Central 认证失败

**检查**:
- ✅ `MAVEN_USERNAME` 和 `MAVEN_PASSWORD` 是否正确
- ✅ Token 是否过期
- ✅ 是否有发布权限

### 问题 4: 版本已存在

**解决**:
- Maven Central 不允许覆盖已发布的版本
- 需要修改为新的版本号

---

## 📝 最佳实践

### 1. 版本号管理

使用语义化版本号:
- `3.0.0` → `3.0.1` (Bug 修复)
- `3.0.1` → `3.1.0` (新功能)
- `3.1.0` → `4.0.0` (破坏性变更)

### 2. 开发流程

```bash
# 开发阶段: 使用 SNAPSHOT 版本
<version>3.1.0-SNAPSHOT</version>

# 发布阶段: 移除 SNAPSHOT
<version>3.1.0</version>

# 发布后: 准备下一个开发版本
<version>3.1.1-SNAPSHOT</version>
```

### 3. 提交信息

使用规范的提交信息:
```bash
git commit -m "chore: bump version to 3.0.1"
git commit -m "chore: release version 3.0.1"
git commit -m "chore: prepare for next development iteration"
```

---

## 🔒 安全建议

- ✅ 定期更新 Maven Central Token
- ✅ 定期更新 GPG 密钥
- ✅ 不要在代码中硬编码凭证
- ✅ 使用 GitHub Secrets 管理敏感信息
- ✅ 不要将 `.github/secrets/` 目录提交到仓库

---

## 📚 参考资料

- [Maven Central Publishing Guide](https://central.sonatype.org/publish/publish-guide/)
- [GitHub Actions Documentation](https://docs.github.com/en/actions)
- [GPG Documentation](https://gnupg.org/documentation/)
- [Semantic Versioning](https://semver.org/)

---

## 🆘 需要帮助？

- 查看 [详细配置指南](SETUP_GUIDE.md)
- 查看 [GitHub Actions 日志](https://github.com/chenqi92/allbs-excel/actions)
- 提交 [Issue](https://github.com/chenqi92/allbs-excel/issues)
- 联系维护者: chenqi92104@icloud.com

