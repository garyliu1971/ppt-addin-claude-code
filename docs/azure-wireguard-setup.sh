# ============================================================
# Azure VM + WireGuard 分流模式搭建指南
# 目标：只让 claude.ai / anthropic 走 VPN，其余流量直连
#
# 架构：
#   浏览器 ─→ SwitchyOmega ─┬─→ 本地直连 (99% 流量)
#                            └─→ SOCKS5 :1080 ─→ WireGuard 隧道 ─→ VM tinyproxy ─→ Internet
#
# VM 需要开放端口：22 (SSH), 51820 (UDP, WireGuard)
# ============================================================

# ━━━ 第一步：Azure Portal 上创建 VM ━━━

# 1. 打开 https://portal.azure.com → 创建资源 → Virtual Machine
# 2. 关键配置：
#    Region:        East Asia (香港) 或 Southeast Asia (新加坡)
#    Image:         Ubuntu Server 22.04 LTS
#    Size:          Standard_B1s (1 vCPU, 1GB RAM, ~$7/月)
#    Auth type:     SSH public key (生成或用已有的)
#    Public IP:     必须开启（动态即可）
# 3. 在 Networking 标签页，添加 Inbound rules：
#    - SSH (22)          ← 默认已加
#    - UDP 51820         ← WireGuard，手动添加
# 4. 创建完成 → 记下 Public IP

# ━━━ 第二步：SSH 进 VM，运行下面脚本 ━━━

# 本地 PowerShell：
# ssh -i your-key.pem azureuser@<VM_PUBLIC_IP>
# 然后复制粘贴下面整个脚本
```

```bash
#!/bin/bash
# ============================================================
# setup-split-tunnel.sh
# WireGuard + tinyproxy 分流模式，一键搭建
# ============================================================
set -e

echo "═══════════════════════════════════════════"
echo "  WireGuard + tinyproxy 分流模式 安装"
echo "═══════════════════════════════════════════"

# ── 更新 ──
sudo apt update && sudo apt upgrade -y

# ── 1) 安装 WireGuard ──
sudo apt install -y wireguard

SERVER_PRIV=$(wg genkey)
SERVER_PUB=$(echo "$SERVER_PRIV" | wg pubkey)
CLIENT_PRIV=$(wg genkey)
CLIENT_PUB=$(echo "$CLIENT_PRIV" | wg pubkey)

SERVER_IP="10.0.0.1"
CLIENT_IP="10.0.0.2"
PORT=51820
PROXY_PORT=8888
PUBLIC_IP=$(curl -s ifconfig.me)

echo ""
echo "  公网 IP:    $PUBLIC_IP"
echo "  WG 端口:    $PORT"
echo "  Proxy 端口: $PROXY_PORT"
echo ""

# ── 服务器 WireGuard 配置 ──
sudo tee /etc/wireguard/wg0.conf > /dev/null <<EOF
[Interface]
Address = $SERVER_IP/24
ListenPort = $PORT
PrivateKey = $SERVER_PRIV
PostUp = iptables -A FORWARD -i wg0 -j ACCEPT
PostUp = iptables -t nat -A POSTROUTING -o eth0 -j MASQUERADE
PostDown = iptables -D FORWARD -i wg0 -j ACCEPT
PostDown = iptables -t nat -D POSTROUTING -o eth0 -j MASQUERADE

[Peer]
PublicKey = $CLIENT_PUB
AllowedIPs = $CLIENT_IP/32
EOF

sudo sed -i 's/#net.ipv4.ip_forward=1/net.ipv4.ip_forward=1/' /etc/sysctl.conf
sudo sysctl -p

sudo systemctl enable wg-quick@wg0
sudo systemctl start wg-quick@wg0

# ── 2) 安装 tinyproxy (轻量 HTTP 代理) ──
sudo apt install -y tinyproxy

sudo tee /etc/tinyproxy/tinyproxy.conf > /dev/null <<EOF
User tinyproxy
Group tinyproxy
Port $PROXY_PORT
Listen $SERVER_IP
Timeout 600
DefaultErrorFile "/usr/share/tinyproxy/default.html"
StatFile "/usr/share/tinyproxy/stats.html"
LogLevel Info
MaxClients 100
MinSpareServers 2
MaxSpareServers 5
StartServers 3
MaxRequestsPerChild 0
Allow 127.0.0.1
Allow $CLIENT_IP
ViaProxyName "tinyproxy"
ConnectPort 443
ConnectPort 80
EOF

sudo systemctl enable tinyproxy
sudo systemctl restart tinyproxy

# ── 3) 写入 DNS 分流配置（可选：让 VM 解析更准） ──
sudo tee -a /etc/hosts > /dev/null <<EOF
# (分流模式不需要改 hosts，tinyproxy 直接代理 HTTP 请求)
EOF

echo ""
echo "═══════════════════════════════════════════"
echo "  ✅ 安装完成！"
echo "═══════════════════════════════════════════"
echo ""

# ── 客户端配置 ──
cat > ~/client-wg-split.conf <<EOF
[Interface]
PrivateKey = $CLIENT_PRIV
Address = $CLIENT_IP/24
# ⚠️ 分流关键：只路由隧道网段，不劫持全局流量
# DNS 走本地，不影响国内网站

[Peer]
PublicKey = $SERVER_PUB
Endpoint = $PUBLIC_IP:$PORT
# ✅ 只允许隧道内网 IP，其余流量走本地
AllowedIPs = 10.0.0.0/24
PersistentKeepalive = 25
EOF

echo "━━━━━━━━━━ 客户端 WireGuard 配置 ━━━━━━━━━━"
cat ~/client-wg-split.conf
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo ""
echo "📥 配置文件: ~/client-wg-split.conf"
echo ""
echo "📋 下一步："
echo "  1. 把 client-wg-split.conf 导入 Windows WireGuard 客户端"
echo "  2. 连接 WireGuard"
echo "  3. 测试: ping 10.0.0.1 (应通)"
echo "  4. 浏览器装 SwitchyOmega 扩展"
echo "  5. 新建代理: SOCKS5 → 10.0.0.1:$PROXY_PORT"
echo "  6. 新建规则: *.claude.ai → 走代理; *.anthropic.com → 走代理"
echo "  7. 其余全部直连"
echo ""
echo "🔍 验证: curl --proxy http://10.0.0.1:$PROXY_PORT https://claude.ai"

