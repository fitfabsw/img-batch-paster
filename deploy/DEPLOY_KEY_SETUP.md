# GitHub Deploy Key (lab mini)

Deploy script 在 lab mini 上 `git pull` 時需要對 GitHub 認證。用 **deploy key**（repo-scoped、無 passphrase）最乾淨；public / private repo 都適用。

## 1. 產 key（在 lab mini 上）

```bash
ssh-keygen -t ed25519 \
  -f ~/.ssh/github_<reponame>_key \
  -N "" \
  -C "lab-mini deploy: <reponame>"
```

| 參數         | 用途                                  |
| ------------ | ------------------------------------- |
| `-t ed25519` | 演算法（比 RSA 短、快，現代推薦）     |
| `-f <path>`  | 私鑰位置；`.pub` 同名自動產出         |
| `-N ""`      | 空 passphrase（無人值守 deploy 必須） |
| `-C "..."`   | comment，方便 GitHub 端識別           |

## 2. 註冊到 GitHub

```bash
cat ~/.ssh/github_<reponame>_key.pub
```

複製整行 → `https://github.com/<owner>/<reponame>/settings/keys` → **Add deploy key**：

- **Title**: `lab-mini`
- **Key**: 貼上 `.pub` 內容
- **Allow write access**: 不勾（只 pull）

## 3. SSH config

### 3a. 單 repo 模式（lab mini 只 deploy 一個 repo）

`~/.ssh/config` 直接綁 `github.com`：

```
Host github.com
    HostName github.com
    User git
    IdentityFile ~/.ssh/github_<reponame>_key
    IdentitiesOnly yes
```

Repo remote URL 維持標準形式：`git@github.com:<owner>/<reponame>.git`，不用改。

### 3b. 多 repo 模式（lab mini 同時 deploy 兩個以上 repo）

兩把 key 都綁 `github.com` 會打架 — SSH 只看 hostname 挑 key，無法分辨哪個 repo 該用哪把，會試錯 → `Permission denied`。

解法：**Host alias**。每個 repo 各自一段 config：

```
Host github-imgbatch
    HostName github.com
    User git
    IdentityFile ~/.ssh/github_imgbatch_key
    IdentitiesOnly yes

Host github-otherrepo
    HostName github.com
    User git
    IdentityFile ~/.ssh/github_otherrepo_key
    IdentitiesOnly yes
```

Repo remote URL 改用 alias（不再是 `github.com`）：

```bash
# 新 clone
git clone git@github-imgbatch:fitfabsw/img-batch-paster.git

# 既有 repo 改 remote
cd ~/img-batch-paster
git remote set-url origin git@github-imgbatch:fitfabsw/img-batch-paster.git
```

`HostName github.com` 那行是 SSH 真實要連的目的地；前面的 `Host github-imgbatch` 只是讓 SSH 看到這個 alias 時知道「套用這條規則 + 用這把 key」。

## 4. 測試

單 repo（3a）：

```bash
ssh -T git@github.com
# → Hi <owner>/<reponame>! You've successfully authenticated...
```

多 repo（3b）：

```bash
ssh -T git@github-imgbatch
# → Hi fitfabsw/img-batch-paster! ...
```

通過後 deploy script 即可正常運作。

---

## 常見錯誤排解

### `Permission denied (publickey)`

- 確認 GitHub repo Deploy keys 已新增 `.pub`
- `ssh-add -l` 看 agent 有沒有載 key（沒載也沒關係，SSH config 指定的 IdentityFile 會優先）
- `ssh -vT git@github.com`（或 alias）看 verbose log，注意 `Offering public key:` 行有沒有指到對的檔案

### `Key is already in use`

GitHub 不允許同一把 deploy key 用在多個 repo。每個 repo 各自產一把。

### 重開機後 deploy 突然失敗 + 跳 passphrase prompt

代表 key 設了 passphrase + ssh-agent 重開後沒持久化。解法：產一把無 passphrase 的 deploy key（如本文件 step 1），別依賴 agent 暫存。

### 想撤銷

GitHub repo Deploy keys 頁面點該 key 旁的刪除按鈕；同時在 lab mini `rm ~/.ssh/github_<reponame>_key*` 清掉檔案。
