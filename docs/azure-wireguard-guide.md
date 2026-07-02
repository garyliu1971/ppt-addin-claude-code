# Azure + WireGuard 分流模式 完整指南

## 架构（分流）

```
浏览器 ── SwitchyOmega ─┬─→ 本地直连 (99% 流量：百度、B站、微信...)
                         └─→ SOCKS5 10.0.0.1:8888 ─→ WireGuard 隧道 ─→ Azure VM tinyproxy ─→ claude.ai
```

**只有 claude.ai / anthropic.com 走 VPN，其余流量不受影响。**

---

## 第一步：Azure 创建 VM

1. 打开 [Azure Portal](https://portal.azure.com)
2. **创建资源** → **Virtual Machine**
3. 关键配置：

| 设置 | 值 |
|---|---|
| Region | **East Asia**（香港）或 **Southeast Asia**（新加坡） |
| Image | Ubuntu Server 22.04 LTS |
| Size | **Standard_B1s**（~$7/月，1 vCPU / 1 GB） |
| Auth type | SSH public key |
| Public IP | 必须开启 |

4. **Networking** 标签 → 添加入站规则：

| 端口 | 协议 | 用途 |
|---|---|---|
| 22 | TCP | SSH |
| 51820 | UDP | WireGuard |

5. 创建 → 记下 **Public IP**

---

## 第二步：SSH 登入 VM + 运行脚本

```powershell
ssh -i your-key.pem azureuser@<VM_PUBLIC_IP>
```

把 `docs/azure-wireguard-setup.sh` 的内容粘贴到终端执行，脚本自动完成：
- 安装 WireGuard + 配置
- 安装 tinyproxy（HTTP 代理，监听 `10.0.0.1:8888`）
- 输出客户端配置文件

---

## 第三步：Windows 配置 WireGuard（分流！）

1. 下载 [WireGuard for Windows](https://www.wireguard.com/install/)
2. 安装 → **Add Tunnel** → **Add empty tunnel...**
3. 粘贴 VM 输出的 `client-wg-split.conf`：

```ini
[Interface]
PrivateKey = <客户端私钥>
Address = 10.0.0.2/24

[Peer]
PublicKey = <服务器公钥>
Endpoint = <VM_PUBLIC_IP>:51820
# ⚠️ 分流关键：只路由隧道网段，不劫持全局流量
AllowedIPs = 10.0.0.0/24
PersistentKeepalive = 25
```

4. 点击 **Activate**，验证：`ping 10.0.0.1` 应通

---

## 第四步：浏览器配置域名分流（SwitchyOmega）

1. Chrome/Edge 安装 [SwitchyOmega](https://chrome.google.com/webstore/detail/proxy-switchyomega/padekgcemlokbadohgkifijomclgjgif)

2. 新建代理配置：
   - 点击 SwitchyOmega 图标 → **Options**
   - **New profile** → 命名为 `Claude VPN`
   - Protocol: **HTTP** | Server: `10.0.0.1` | Port: `8888`

3. 新建自动切换规则（auto switch）：
   - **Condition type**: `Host wildcard`
   - 添加规则：

   | 条件 | 代理配置 |
   |---|---|
   | `*.claude.ai` | Claude VPN |
   | `*.anthropic.com` | Claude VPN |
   | `claude.ai` | Claude VPN |
   | `anthropic.com` | Claude VPN |

   - **Default**: `[Direct]`（其余全部直连）

4. 浏览器地址栏旁选 **auto switch** 模式

---

## 验证

```powershell
# 1) 验证 WireGuard 隧道通
ping 10.0.0.1

# 2) 验证代理可用
curl --proxy http://10.0.0.1:8888 https://claude.ai -I

# 3) 验证本地 IP 没被劫持（应该是国内 IP，不是 Azure 的）
curl ifconfig.me
```

---

## 月成本估算

| 资源 | 月费（约） |
|---|---|
| B1s VM (1 vCPU / 1 GB) | $7 |
| 30 GB Standard SSD | $3 |
| 动态 Public IP | $3 |
| 流量（仅 claude.ai，极少） | <$1 |
| **合计** | **~$14/月** |

> 不用时关机省 VM 费（IP + Disk 约 $6/月照收）

---

## 常见问题

**Q: 非 Chrome 浏览器怎么分流？**
A: Firefox 用 **FoxyProxy**，同样支持域名规则；或者用 **Proxifier**（系统级，所有程序都能域名分流，收费 ~$40）。

**Q: 想让 VS Code / Terminal 也走代理？**
```powershell
$env:HTTP_PROXY = "http://10.0.0.1:8888"
$env:HTTPS_PROXY = "http://10.0.0.1:8888"
```

**Q: Cloudflare 拦怎么办？**
A: Azure 香港/新加坡 IP 相对干净。如遇到验证码，在 VM 上换 `tinyproxy` 端口或尝试 `squid` 替代。

---

## 月成本估算

| 资源 | 规格 | 月费（约） |
|---|---|---|
| VM B1s | 1 vCPU / 1 GB | $7 |
| Managed Disk | 30 GB Standard SSD | $3 |
| Public IP | 动态 | $3 |
| 流量 | 1 TB 出站 | $0.08/GB |
| **合计** | | **~$15/月** |

> 不用时关机节省 VM 费用，Public IP 和 Disk 照常收费。

## 常见问题

**Q: Cloudflare 拦了怎么办？**
A: Azure 香港/新加坡的 IP 相对干净，但仍可能被 CF 挑战。如果遇到验证码，可以：
- 换一个 VM 区域试试
- 用 API 端点（`api.anthropic.com`）代替网页版
- 在 VM 上装 Firefox 走代理模式

**Q: 怎么节省费用？**
A: 用 **Spot VM**（竞价实例），B1s 可以压到 ~$2/月，但随时可能被回收。
