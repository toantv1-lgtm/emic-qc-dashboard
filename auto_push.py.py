import base64
import os
import sys
import time
import requests

# --- CẤU HÌNH THÔNG TIN GITHUB ---
GITHUB_USER = "TÊN_TÀI_KHOẢN_GITHUB_CỦA_BẠN"
REPO_NAME = "emic-qc-dashboard"
GITHUB_TOKEN = "MÃ_TOKEN_CỦA_BẠN"  # Mã ghp_... lấy ở Bước 1.2
FILE_TO_SYNC = "Report_Database.db"


def push_db_to_github():
  if not os.path.exists(FILE_TO_SYNC):
    print(f"❌ Không tìm thấy file {FILE_TO_SYNC}!")
    return False

  url = f"https://api.github.com/repos/{GITHUB_USER}/{REPO_NAME}/contents/{FILE_TO_SYNC}"
  headers = {
      "Authorization": f"Bearer {GITHUB_TOKEN}",
      "Accept": "application/vnd.github.v3+json",
  }

  print(f"⏳ Đang tải file {FILE_TO_SYNC} lên GitHub...")

  try:
    # 1. Lấy mã SHA hiện tại của file trên GitHub (nếu có)
    sha = None
    res = requests.get(url, headers=headers)
    if res.status_code == 200:
      sha = res.json().get("sha")

    # 2. Đọc và mã hóa file SQLite
    with open(FILE_TO_SYNC, "rb") as f:
      file_content = base64.b64encode(f.read()).decode("utf-8")

    # 3. Gửi Request đè dữ liệu lên GitHub
    payload = {
        "message": (
            "Auto-update database from Local:"
            f" {time.strftime('%Y-%m-%d %H:%M:%S')}"
        ),
        "content": file_content,
    }
    if sha:
      payload["sha"] = sha

    put_res = requests.put(url, json=payload, headers=headers)
    if put_res.status_code in [200, 201]:
      print("✅ ĐÃ ĐỒNG BỘ THÀNH CÔNG LÊN CLOUD!")
      return True
    else:
      print("❌ Lỗi đồng bộ:", put_res.json())
      return False

  except Exception as e:
    print(f"❌ Lỗi kết nối: {e}")
    return False


# --- THEO DÕI TỰ ĐỘNG KHI FILE THAY ĐỔI ---
if __name__ == "__main__":
  print("🚀 Đang khởi chạy tiến trình tự động theo dõi file Database...")
  last_mtime = 0

  while True:
    if os.path.exists(FILE_TO_SYNC):
      current_mtime = os.path.getmtime(FILE_TO_SYNC)
      if current_mtime != last_mtime:
        if last_mtime != 0:
          print(f"\n🔄 Phát hiện file {FILE_TO_SYNC} đã được cập nhật mới!")
          push_db_to_github()
        else:
          # Tải lên lần đầu khi mới bật script
          push_db_to_github()
        last_mtime = current_mtime
    time.sleep(5)  # Kiểm tra sự thay đổi mỗi 5 giây