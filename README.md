# 大挑仓库

团队竞赛项目的共享仓库：**代码**放在 `大挑/`，**参考文献**放在 `论文/`。

## 目录结构

```
datiaocangku/
├── .gitattributes   # 换行符与二进制文件规则（PDF 等不做文本转换）
├── .gitignore       # 忽略规则
├── 大挑/            # 项目代码
│   └── lalala.py
└── 论文/            # 参考文献（PDF）
```

## 环境要求

- Git 2.40 或更高
- Python 3.10 或更高（`大挑/` 下的代码使用）

## 获取代码

```powershell
git clone https://github.com/yhc6661/datiaocangku.git
cd datiaocangku
```

国内网络直连 GitHub 经常出现 `Recv failure: Connection was reset`，需要给 git 单独配代理
（git **不会**自动使用 Windows 的系统代理设置）：

```powershell
git config --global http.proxy  http://127.0.0.1:7897
git config --global https.proxy http://127.0.0.1:7897
```

端口 `7897` 是代理客户端的混合端口，按自己机器上的实际端口改。取消代理：

```powershell
git config --global --unset http.proxy
git config --global --unset https.proxy
```

## 协作流程

尽量**不要直接往 `main` 推**，按分支 + Pull Request 走，避免互相覆盖。

```powershell
git pull                          # 1. 开工前同步
git checkout -b li0808/功能名      # 2. 建自己的分支，名字带自己的前缀
# ... 写代码 ...
git add .
git commit -m "说明了什么改动"      # 3. 提交
git push -u origin li0808/功能名   # 4. 推到远程
```

推完之后在 GitHub 页面上点 `Compare & pull request` 提交 PR，由仓库所有者合并。

### 提交信息前缀约定（建议）

| 前缀 | 用于 |
| --- | --- |
| `feat:` | 新增功能 |
| `fix:` | 修复问题 |
| `docs:` | 文档改动 |
| `data:` | 文献、数据集 |
| `chore:` | 杂项（配置、忽略规则等） |

## 注意事项

- `论文/` 下的 PDF 已在 `.gitattributes` 里标记为 `binary`，不会被换行符转换破坏。
- PDF 属于二进制大文件，仓库目前约 6 MB。如果后续文献积累到几百 MB，建议改用
  [Git LFS](https://git-lfs.com/) 管理，否则克隆会越来越慢。
- 仓库当前为**公开**状态，`论文/` 里是正式发表的期刊论文，公开托管涉及版权，
  建议在 GitHub 上改为 Private（`Settings` → `General` → 最底部 `Danger Zone`）。

## 成员

| GitHub 账号 | 角色 |
| --- | --- |
| [yhc6661](https://github.com/yhc6661) | 仓库所有者 |
| [li0808-afk](https://github.com/li0808-afk) | 协作者 |

## 待补充

- [ ] 项目简介 / 研究题目
- [ ] 代码运行方式（依赖、入口）
- [ ] 数据与文献来源说明
