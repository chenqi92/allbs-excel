#!/bin/bash

# GPG 密钥生成和配置脚本
# 用于快速生成 GPG 密钥并导出为 GitHub Secrets 所需的格式

set -e

echo "=========================================="
echo "  GPG 密钥生成和配置脚本"
echo "=========================================="
echo ""

# 颜色定义
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

# 检查 GPG 是否安装
if ! command -v gpg &> /dev/null; then
    echo -e "${RED}❌ GPG 未安装${NC}"
    echo ""
    echo "请先安装 GPG:"
    echo "  macOS:   brew install gnupg"
    echo "  Ubuntu:  sudo apt-get install gnupg"
    echo "  CentOS:  sudo yum install gnupg"
    exit 1
fi

echo -e "${GREEN}✅ GPG 已安装${NC}"
echo ""

# 询问是否生成新密钥
echo "是否需要生成新的 GPG 密钥？"
echo "  1) 是，生成新密钥"
echo "  2) 否，使用现有密钥"
read -p "请选择 [1/2]: " choice

if [ "$choice" = "1" ]; then
    echo ""
    echo "=========================================="
    echo "  生成新的 GPG 密钥"
    echo "=========================================="
    echo ""
    
    # 收集信息
    read -p "请输入你的姓名: " name
    read -p "请输入你的邮箱: " email
    read -sp "请输入 GPG 密钥密码 (至少 8 位): " passphrase
    echo ""
    read -sp "请再次输入密码: " passphrase2
    echo ""
    
    if [ "$passphrase" != "$passphrase2" ]; then
        echo -e "${RED}❌ 两次密码不一致${NC}"
        exit 1
    fi
    
    # 生成密钥配置文件
    cat > /tmp/gpg-gen-key.conf << EOF
%echo Generating GPG key
Key-Type: RSA
Key-Length: 4096
Subkey-Type: RSA
Subkey-Length: 4096
Name-Real: $name
Name-Email: $email
Expire-Date: 0
Passphrase: $passphrase
%commit
%echo Done
EOF
    
    echo ""
    echo "正在生成 GPG 密钥，这可能需要几分钟..."
    gpg --batch --gen-key /tmp/gpg-gen-key.conf
    rm /tmp/gpg-gen-key.conf
    
    # 获取密钥 ID
    KEY_ID=$(gpg --list-secret-keys --keyid-format LONG "$email" | grep sec | awk '{print $2}' | cut -d'/' -f2)
    
    echo -e "${GREEN}✅ GPG 密钥生成成功${NC}"
    echo "密钥 ID: $KEY_ID"
    
else
    echo ""
    echo "=========================================="
    echo "  现有的 GPG 密钥"
    echo "=========================================="
    echo ""
    
    # 列出现有密钥
    gpg --list-secret-keys --keyid-format LONG
    
    echo ""
    read -p "请输入要使用的密钥 ID (例如: ABCD1234EFGH5678): " KEY_ID
    read -sp "请输入该密钥的密码: " passphrase
    echo ""
fi

echo ""
echo "=========================================="
echo "  导出密钥"
echo "=========================================="
echo ""

# 导出私钥 (Base64 编码)
echo "正在导出私钥..."
GPG_PRIVATE_KEY=$(gpg --armor --export-secret-keys "$KEY_ID" | base64)

# 保存到文件
mkdir -p .github/secrets
echo "$GPG_PRIVATE_KEY" > .github/secrets/gpg-private-key.txt
echo "$passphrase" > .github/secrets/gpg-passphrase.txt

echo -e "${GREEN}✅ 密钥已导出${NC}"
echo ""

# 上传公钥到密钥服务器
echo "=========================================="
echo "  上传公钥到密钥服务器"
echo "=========================================="
echo ""

read -p "是否上传公钥到密钥服务器？[y/N]: " upload_choice

if [ "$upload_choice" = "y" ] || [ "$upload_choice" = "Y" ]; then
    echo "正在上传公钥..."
    gpg --keyserver keyserver.ubuntu.com --send-keys "$KEY_ID" || echo -e "${YELLOW}⚠️  上传到 keyserver.ubuntu.com 失败${NC}"
    gpg --keyserver keys.openpgp.org --send-keys "$KEY_ID" || echo -e "${YELLOW}⚠️  上传到 keys.openpgp.org 失败${NC}"
    echo -e "${GREEN}✅ 公钥上传完成${NC}"
fi

echo ""
echo "=========================================="
echo "  配置 GitHub Secrets"
echo "=========================================="
echo ""

echo "请在 GitHub 仓库中配置以下 Secrets:"
echo ""
echo -e "${YELLOW}1. GPG_PRIVATE_KEY${NC}"
echo "   值: (已保存到 .github/secrets/gpg-private-key.txt)"
echo ""
echo -e "${YELLOW}2. GPG_PASSPHRASE${NC}"
echo "   值: (已保存到 .github/secrets/gpg-passphrase.txt)"
echo ""
echo -e "${YELLOW}3. MAVEN_USERNAME${NC}"
echo "   值: 从 https://central.sonatype.com/ 获取"
echo ""
echo -e "${YELLOW}4. MAVEN_PASSWORD${NC}"
echo "   值: 从 https://central.sonatype.com/ 获取"
echo ""

echo "=========================================="
echo "  配置步骤"
echo "=========================================="
echo ""
echo "1. 打开 GitHub 仓库"
echo "2. 进入 Settings → Secrets and variables → Actions"
echo "3. 点击 'New repository secret'"
echo "4. 添加上述 4 个 Secrets"
echo ""

echo "=========================================="
echo "  密钥信息"
echo "=========================================="
echo ""
echo "密钥 ID: $KEY_ID"
echo "私钥文件: .github/secrets/gpg-private-key.txt"
echo "密码文件: .github/secrets/gpg-passphrase.txt"
echo ""

echo -e "${GREEN}✅ 配置完成！${NC}"
echo ""
echo -e "${RED}⚠️  重要提示:${NC}"
echo "1. 请妥善保管 .github/secrets/ 目录中的文件"
echo "2. 不要将这些文件提交到 Git 仓库"
echo "3. 配置完 GitHub Secrets 后，可以删除这些文件"
echo ""

# 添加到 .gitignore
if ! grep -q ".github/secrets/" .gitignore 2>/dev/null; then
    echo ".github/secrets/" >> .gitignore
    echo -e "${GREEN}✅ 已添加 .github/secrets/ 到 .gitignore${NC}"
fi

