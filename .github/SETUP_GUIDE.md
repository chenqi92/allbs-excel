# GitHub Actions 自动发布配置指南

本指南将帮助你配置 GitHub Actions，实现自动发布到 Maven Central 和 GitHub Packages。

---

## 📋 前置要求

### 1. Maven Central (Sonatype) 账号

你需要在 [Sonatype Central](https://central.sonatype.com/) 注册账号并获取：
- **用户名** (Username)
- **密码** (Password) 或 **Token**

### 2. GPG 密钥

用于签名 Maven 构建产物。

### 3. GitHub 账号

需要有仓库的管理员权限。

---

## 🔧 配置步骤

### 步骤 1: 生成 GPG 密钥

如果你还没有 GPG 密钥，需要先生成：

```bash
# 1. 生成 GPG 密钥对
gpg --full-generate-key

# 选择：
# - 密钥类型: RSA and RSA (默认)
# - 密钥长度: 4096
# - 有效期: 0 (永不过期) 或根据需要设置
# - 真实姓名: 你的名字
# - 电子邮件: 你的邮箱 (建议使用 GitHub 邮箱)
# - 注释: 可选
# - 密码: 设置一个强密码 (记住这个密码，后面需要用到)

# 2. 查看生成的密钥
gpg --list-secret-keys --keyid-format LONG

# 输出示例:
# sec   rsa4096/ABCD1234EFGH5678 2024-01-01 [SC]
#       1234567890ABCDEF1234567890ABCDEF12345678
# uid                 [ultimate] Your Name <your.email@example.com>
# ssb   rsa4096/1234567890ABCDEF 2024-01-01 [E]

# 记住密钥 ID: ABCD1234EFGH5678

# 3. 导出私钥 (Base64 编码)
gpg --armor --export-secret-keys ABCD1234EFGH5678 | base64 > gpg-private-key.txt

# 4. 上传公钥到密钥服务器
gpg --keyserver keyserver.ubuntu.com --send-keys ABCD1234EFGH5678
gpg --keyserver keys.openpgp.org --send-keys ABCD1234EFGH5678
```

---

### 步骤 2: 配置 GitHub Secrets

在 GitHub 仓库中配置以下 Secrets：

1. **进入仓库设置**
   - 打开你的 GitHub 仓库
   - 点击 `Settings` → `Secrets and variables` → `Actions`
   - 点击 `New repository secret`

2. **添加以下 Secrets**:

| Secret 名称 | 说明 | 获取方式 |
|------------|------|---------|
| `MAVEN_USERNAME` | Maven Central 用户名 | 从 [Sonatype Central](https://central.sonatype.com/) 获取 |
| `MAVEN_PASSWORD` | Maven Central 密码/Token | 从 Sonatype Central 生成 Token |
| `GPG_PRIVATE_KEY` | GPG 私钥 (Base64) | 从 `gpg-private-key.txt` 文件复制 |
| `GPG_PASSPHRASE` | GPG 密钥密码 | 生成 GPG 密钥时设置的密码 |

**详细步骤**:

#### 2.1 获取 Maven Central 凭证

1. 访问 [https://central.sonatype.com/](https://central.sonatype.com/)
2. 登录你的账号
3. 点击右上角头像 → `View Account`
4. 点击 `Generate User Token`
5. 复制生成的 **Username** 和 **Password**
6. 在 GitHub Secrets 中添加：
   - `MAVEN_USERNAME`: 粘贴 Username
   - `MAVEN_PASSWORD`: 粘贴 Password

#### 2.2 添加 GPG 密钥

1. 打开 `gpg-private-key.txt` 文件
2. 复制全部内容（包括 Base64 编码的字符串）
3. 在 GitHub Secrets 中添加：
   - `GPG_PRIVATE_KEY`: 粘贴复制的内容
   - `GPG_PASSPHRASE`: 输入你生成 GPG 密钥时设置的密码

---

### 步骤 3: 验证配置

#### 3.1 检查 Secrets

确保以下 4 个 Secrets 都已正确配置：

```
✅ MAVEN_USERNAME
✅ MAVEN_PASSWORD
✅ GPG_PRIVATE_KEY
✅ GPG_PASSPHRASE
```

#### 3.2 测试 Workflow

1. **手动触发测试**:
   - 进入 `Actions` 标签页
   - 选择 `Publish on Version Change` workflow
   - 点击 `Run workflow`
   - 选择 `main` 分支
   - 点击 `Run workflow` 按钮

2. **修改版本号触发**:
   ```bash
   # 修改 pom.xml 中的版本号
   # 例如: 从 3.0.0 改为 3.0.1
   
   git add pom.xml
   git commit -m "chore: bump version to 3.0.1"
   git push origin main
   ```

3. **查看执行结果**:
   - 进入 `Actions` 标签页
   - 查看最新的 workflow 运行状态
   - 如果成功，会看到绿色的 ✅
   - 如果失败，点击查看日志排查问题

---

## 📦 发布流程

### 自动发布流程

1. **修改版本号**:
   ```xml
   <!-- pom.xml -->
   <version>3.0.1</version>  <!-- 从 3.0.0 改为 3.0.1 -->
   ```

2. **提交并推送**:
   ```bash
   git add pom.xml
   git commit -m "chore: bump version to 3.0.1"
   git push origin main
   ```

3. **自动执行**:
   - GitHub Actions 检测到 `pom.xml` 变化
   - 检查版本号是否改变
   - 如果版本号改变，自动执行：
     - ✅ 编译项目
     - ✅ 运行测试
     - ✅ 发布到 Maven Central
     - ✅ 发布到 GitHub Packages
     - ✅ 创建 GitHub Release (仅 Release 版本)
     - ✅ 创建 Git Tag (仅 Release 版本)

4. **查看结果**:
   - Maven Central: https://central.sonatype.com/artifact/cn.allbs/allbs-excel
   - GitHub Packages: https://github.com/chenqi92/allbs-excel/packages
   - GitHub Releases: https://github.com/chenqi92/allbs-excel/releases

---

## 🔍 版本类型

### Release 版本

```xml
<version>3.0.0</version>
```

- 发布到 Maven Central
- 发布到 GitHub Packages
- 创建 GitHub Release
- 创建 Git Tag

### SNAPSHOT 版本

```xml
<version>3.0.1-SNAPSHOT</version>
```

- 发布到 Maven Central (SNAPSHOT 仓库)
- 发布到 GitHub Packages
- **不会**创建 GitHub Release
- **不会**创建 Git Tag

---

## 🛠️ 故障排查

### 问题 1: GPG 签名失败

**错误信息**: `gpg: signing failed: No secret key`

**解决方案**:
1. 检查 `GPG_PRIVATE_KEY` 是否正确 Base64 编码
2. 检查 `GPG_PASSPHRASE` 是否正确
3. 重新生成并导出 GPG 密钥

### 问题 2: Maven Central 认证失败

**错误信息**: `401 Unauthorized`

**解决方案**:
1. 检查 `MAVEN_USERNAME` 和 `MAVEN_PASSWORD` 是否正确
2. 确认 Token 是否过期
3. 重新生成 User Token

### 问题 3: 版本已存在

**错误信息**: `Version already exists`

**解决方案**:
1. Maven Central 不允许覆盖已发布的版本
2. 需要修改为新的版本号
3. 建议使用语义化版本号 (Semantic Versioning)

### 问题 4: Workflow 没有触发

**可能原因**:
1. 只修改了 `pom.xml` 以外的文件
2. 版本号没有实际改变
3. 推送到了非 `main` 分支

**解决方案**:
1. 确保修改了 `pom.xml` 中的 `<version>` 标签
2. 确保版本号确实发生了变化
3. 推送到 `main` 分支

---

## 📝 最佳实践

### 1. 版本号管理

使用语义化版本号 (Semantic Versioning):
- **主版本号**: 不兼容的 API 修改
- **次版本号**: 向下兼容的功能性新增
- **修订号**: 向下兼容的问题修正

示例:
- `3.0.0` → `3.0.1` (Bug 修复)
- `3.0.1` → `3.1.0` (新功能)
- `3.1.0` → `4.0.0` (破坏性变更)

### 2. 发布前检查

在修改版本号前，确保：
- ✅ 所有测试通过
- ✅ 代码已经过 Code Review
- ✅ 更新了 CHANGELOG
- ✅ 更新了 README (如有必要)

### 3. SNAPSHOT 版本

开发过程中使用 SNAPSHOT 版本:
```xml
<version>3.1.0-SNAPSHOT</version>
```

发布正式版本时移除 `-SNAPSHOT`:
```xml
<version>3.1.0</version>
```

### 4. 安全建议

- ✅ 定期更新 Maven Central Token
- ✅ 定期更新 GPG 密钥
- ✅ 不要在代码中硬编码凭证
- ✅ 使用 GitHub Secrets 管理敏感信息

---

## 📚 参考资料

- [Maven Central Publishing Guide](https://central.sonatype.org/publish/publish-guide/)
- [GitHub Actions Documentation](https://docs.github.com/en/actions)
- [GPG Documentation](https://gnupg.org/documentation/)
- [Semantic Versioning](https://semver.org/)

---

## 🆘 需要帮助？

如果遇到问题，可以：
1. 查看 [GitHub Actions 日志](https://github.com/chenqi92/allbs-excel/actions)
2. 提交 [Issue](https://github.com/chenqi92/allbs-excel/issues)
3. 联系维护者: chenqi92104@icloud.com

