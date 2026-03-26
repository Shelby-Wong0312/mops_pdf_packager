"""Push MOPS 下載結果到 GitHub"""
import subprocess
import sys
import os
import datetime


def main():
    # 切到專案根目錄（src/ 的上一層）
    root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    os.chdir(root_dir)

    print("=" * 60)
    print("  Push 到 GitHub")
    print("=" * 60)
    print()

    # 檢查 git
    try:
        subprocess.run(["git", "--version"], capture_output=True, check=True)
    except FileNotFoundError:
        print("[錯誤] 未偵測到 Git，請先安裝: https://git-scm.com/")
        return

    # git add（依照 .gitignore 規則加入所有變更）
    print("正在加入檔案...")
    subprocess.run(["git", "add", "-A"])

    # 檢查有無變更
    result = subprocess.run(["git", "diff", "--cached", "--quiet"])
    if result.returncode == 0:
        print("沒有新的變更需要上傳。")
        return

    # commit
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    msg = f"更新報告 {now}"
    print(f"正在 commit: {msg}")
    r = subprocess.run(["git", "commit", "-m", msg])
    if r.returncode != 0:
        print("[錯誤] Commit 失敗。")
        return

    # push
    print()
    print("正在 push 到 GitHub...")
    r = subprocess.run(["git", "push"])
    if r.returncode == 0:
        print()
        print("成功上傳到 GitHub！")
    else:
        print()
        print("[錯誤] Push 失敗，請確認網路連線及 GitHub 認證設定。")


if __name__ == "__main__":
    try:
        main()
    finally:
        print()
        input("按 Enter 鍵關閉視窗...")
